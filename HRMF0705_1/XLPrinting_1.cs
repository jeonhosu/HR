using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0705
{
    public class XLPrinting_1
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        //private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        //private int mPrintingLineSTART_1 = 1;  //Line // 1page, 2page, 3page
        //private int mPrintingLineSTART_2 = 6;  //Line // 5page

        private int mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치, 복사 행 누적
        private int mIncrementCopyMAX_1 = 183;  // 1page : 61, 2page : 122, 3page : 183 - 복사되어질 행의 범위
        private int mIncrementCopyMAX_2 = 61;   // 5page
        private int mIncrementCopyMAX_3 = 64;   // 5page

        private int mCopyColumnSTART = 1;    //복사되어  진 행 누적 수
        private int mCopyColumnEND = 43;     //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

        private decimal mBASE_COUNT = 0;    //기본인원수.
        private decimal mOLD_COUNT = 0;     //경로인원수.
        private decimal mBIRTH_COUNT = 0;   //출생인원수.
        private decimal mDISABILITY_COUNT = 0;  //장애인인원수.
        private decimal mCHILD_COUNT = 0;   //6세이하인원수.
        private decimal mWOMAN_COUNT = 0;   //부녀세대

        private decimal mINSURE_AMT = 0;    //보험료.
        private decimal mMEDICAL_AMT = 0;   //의료비.
        private decimal mEDU_AMT = 0;       //교육비.
        private decimal mCREDIT_AMT = 0;    //신용카드.
        private decimal mCHECK_CREDIT_AMT = 0;  //직불카드.
        private decimal mACADE_GIRO_AMT = 0;  //학원비지로납부액
        private decimal mCASH_AMT = 0;      //현금영수증.
        private decimal mTRAD_MARKET_AMT = 0; //전통시장사용액
        private decimal mPUBLIC_TRANSIT_AMT = 0; //대중교통사용액
        private decimal mDONAT_AMT = 0;     //기부금.

        // 기타.
        private decimal mINSURE_ETC_AMT = 0;    //보험료.
        private decimal mMEDICAL_ETC_AMT = 0;   //의료비.
        private decimal mEDU_ETC_AMT = 0;       //교육비.
        private decimal mCREDIT_ETC_AMT = 0;    //신용카드.
        private decimal mCHECK_CREDIT_ETC_AMT = 0;  //체크카드.
        private decimal mACADE_GIRO_ETC_AMT = 0;   //학원비지로납부액
        private decimal mTRAD_MARKET_ETC_AMT = 0;   //전통시장사용액
        private decimal mETC_PUBLIC_TRANSIT_AMT = 0; //대중교통사용액
        private decimal mDONAT_ETC_AMT = 0;     //기부금.

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

        #region ----- Export File Name Methods ----

        private string SetExportFileName(string pExportFileName)
        {
            string vExportFileName = string.Empty;

            try
            {
                vExportFileName = pExportFileName;
                vExportFileName = vExportFileName.Replace("/", "_");
                vExportFileName = vExportFileName.Replace("\\", "_");
                vExportFileName = vExportFileName.Replace("*", "_");
                vExportFileName = vExportFileName.Replace("<", "_");
                vExportFileName = vExportFileName.Replace(">", "_");
                vExportFileName = vExportFileName.Replace("|", "_");
                vExportFileName = vExportFileName.Replace("?", "_");
                vExportFileName = vExportFileName.Replace(":", "_");
                vExportFileName = vExportFileName.Replace(" ", "_");
            }
            catch
            {
            }

            return vExportFileName;
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

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[144];
            pXLColumn = new int[144];

            //----[ 1 page ]------------------------------------------------------------------------------------------------------
            pGDColumn[0] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TYPE");        // 거주 구분(거주자1/거주자2)    
            pGDColumn[1] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONALITY_TYPE");     // 내외국인 구분(내국인1/외국인9)
            pGDColumn[2] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FOREIGN_TAX_YN");       // 외국인단일세율적용
            pGDColumn[3] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSEHOLD_TYPE");       // 세대주 구분(세대주1/세대원2)              
            pGDColumn[4] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_KEEP_TYPE");       // 연말정산구분

            pGDColumn[5] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CORP_NAME");            // 법인명(상호)                  
            pGDColumn[6] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRESIDENT_NAME");       // 대표자(성명)                  
            pGDColumn[7] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("VAT_NUMBER");           // 사업자등록번호              
            pGDColumn[8] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ORG_ADDRESS");          // 소재지(주소)

            pGDColumn[9] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NAME");                 // 성명
            pGDColumn[10] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPRE_NUM");           // 주민번호
            pGDColumn[11] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PERSON_ADDRESS");      // 주소  

            //--------------------------------------------------------------------------------------------------------------------
            // I 근무처별 소득 명세
            //--------------------------------------------------------------------------------------------------------------------                      
            pGDColumn[12] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME");      // 주(현)근무처명
            pGDColumn[13] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NAME1");    // 종(전)1근무처명 
            pGDColumn[14] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NAME2");    // 종(전)2근무처명

            pGDColumn[15] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER");     // 주(현)사업자번호 
            pGDColumn[16] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NUM1");     // 종(전)1사업잡번호
            pGDColumn[17] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NUM2");     // 종(전)2사업잡번호 

            pGDColumn[18] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE");         // 주(현)근무기간
            pGDColumn[19] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE1");        // 종(전)1근무기간
            pGDColumn[20] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE2");        // 종(전)2근무기간

            pGDColumn[21] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE");         // 주(현)감면기간 
            pGDColumn[22] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE1");        // 종(전)1감면기간
            pGDColumn[23] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE2");        // 종(전)2감면기간

            pGDColumn[24] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_PAY_TOT_AMT");     // 주(현)급여 
            pGDColumn[25] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PAY_TOTAL_AMT1");      // 종(전)1급여
            pGDColumn[26] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PAY_TOTAL_AMT2");      // 종(전)2급여

            pGDColumn[27] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_BONUS_TOT_AMT");   // 주(현)상여   
            pGDColumn[28] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BONUS_TOTAL_AMT1");    // 종(전)1상여 
            pGDColumn[29] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BONUS_TOTAL_AMT2");    // 종(전)2상여

            pGDColumn[30] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_ADD_BONUS_AMT");   // 주(현)인정상여   
            pGDColumn[31] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_BONUS_AMT1");      // 종(전)1인정상여
            pGDColumn[32] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_BONUS_AMT2");      // 종(전)2인정상여

            pGDColumn[33] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_STOCK_BENE_AMT");  // 주(현)주식매수선택권
            pGDColumn[34] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STOCK_BENE_AMT1");     // 종(전)1주식매수선택권
            pGDColumn[35] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STOCK_BENE_AMT2");     // 종(전)2주식매수선택권

            pGDColumn[36] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OWNERSHIP_AMT");       // 주(현)우리사주조합인출금
            pGDColumn[37] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OWNERSHIP_AMT1");      // 종(전)1우리사주조합인출금
            pGDColumn[38] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OWNERSHIP_AMT2");      // 종(전)2우리사주조합인출금

            pGDColumn[39] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_TOTAL_AMOUNT");    // 주(현)계     
            pGDColumn[40] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AMOUNT1");       // 종(전)1계   
            pGDColumn[41] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AMOUNT2");       // 종(전)2계  

            //--------------------------------------------------------------------------------------------------------------------
            // II 비과세 및 감면 소득 명세
            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[42] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_OUTSIDE_AMT");  // 비과세_주(현)국외근로
            pGDColumn[43] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OUTSIDE_AMT1");     // 비과세_종(전)1국외근로
            pGDColumn[44] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OUTSIDE_AMT2");     // 비과세_종(전)2국외근로

            pGDColumn[45] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_OT_AMT");       // 비과세_주(현)야간근로수당 
            pGDColumn[46] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OT_AMT1");          // 비과세_종(전)1야간근로수당
            pGDColumn[47] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OT_AMT2");          // 비과세_종(전)2야간근로수당

            pGDColumn[48] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_BIRTH_AMT");    // 비과세_주(현)출산/보육수당
            pGDColumn[49] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_BIRTH_AMT1");       // 비과세_종(전)1출산/보육수당
            pGDColumn[50] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_BIRTH_AMT2");       // 비과세_종(전)2출산/보육수당

            pGDColumn[51] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_FOREIGNER_AMT");// 비과세_주(현)외국인근로자
            pGDColumn[52] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_FOREIGNER_AMT1");   // 비과세_종(전)1외국인근로자
            pGDColumn[53] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_FOREIGNER_AMT2");   // 비과세_종(전)2외국인근로자

            pGDColumn[54] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TOTAL_AMOUNT"); // 비과세_주(현)비과세소득계
            pGDColumn[55] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_TOTAL_AMOUNT1");    // 비과세_종(전)1비과세소득계
            pGDColumn[56] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_TOTAL_AMOUNT2");    // 비과세_종(전)2비과세소득계

            pGDColumn[57] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT"); // 비과세_주(현)감면소득계
            pGDColumn[58] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT1");// 비과세_종(전)1감면소득계
            pGDColumn[59] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT2");// 비과세_종(전)2감면소득계

            //--------------------------------------------------------------------------------------------------------------------
            // III 세액 명세
            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[60] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_IN_TAX_AMT");      // 결정세액_소득세               
            pGDColumn[61] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_LOCAL_TAX_AMT");   // 결정세액_지방소득세               
            pGDColumn[62] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_SP_TAX_AMT");      // 결정세액_농특세               
            pGDColumn[63] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_TAX_AMOUNT");      // 결정세액_계   

            pGDColumn[64] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_COMPANY_NUM1");    // 기납부세액_종(전)1사업자번호  
            pGDColumn[65] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_IN_TAX_AMT1");     // 기납부세액_종(전)1소득세      
            pGDColumn[66] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_LOCAL_TAX_AMT1");  // 기납부세액_종(전)1지방소득세      
            pGDColumn[67] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_SP_TAX_AMT1");     // 기납부세액_종(전)1농특세      
            pGDColumn[68] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_TOTAL_TAX_AMT1");  // 기납부세액_종(전)1계      

            pGDColumn[69] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_COMPANY_NUM2");    // 기납부세액_종(전)2사업자번호  
            pGDColumn[70] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_IN_TAX_AMT2");     // 기납부세액_종(전)2소득세      
            pGDColumn[71] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_LOCAL_TAX_AMT2");  // 기납부세액_종(전)2지방소득세      
            pGDColumn[72] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_SP_TAX_AMT2");     // 기납부세액_종(전)2농특세      
            pGDColumn[73] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_TOTAL_TAX_AMT2");  // 기납부세액_종(전)2계        

            pGDColumn[74] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_IN_TAX_AMT");      // 기납부세액_주(현)소득세       
            pGDColumn[75] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LOCAL_TAX_AMT");   // 기납부세액_주(현)지방소득세       
            pGDColumn[76] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_SP_TAX_AMT");      // 기납부세액_주(현)농특세       
            pGDColumn[77] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_TAX_AMOUNT");      // 기납부세액_주(현)계

            pGDColumn[78] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT");     // 차감징수세액_소득세           
            pGDColumn[79] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT");  // 차감징수세액_지방소득세         
            pGDColumn[80] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_SP_TAX_AMT");     // 차감징수세액_농특세           
            pGDColumn[81] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_TAX_AMOUNT");     // 차감징수세액_계

            //----[ 2 page ]------------------------------------------------------------------------------------------------------
            pGDColumn[82] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_TOT_AMT");           // 총급여
            pGDColumn[83] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PERS_ANNU_BANK_AMT");       // 개인연금저축소득공제

            pGDColumn[84] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_DED_AMT");           // 근로소득공제
            pGDColumn[85] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ANNU_BANK_AMT");            // 연금저축소득공제

            pGDColumn[86] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_AMT");               // 근로소득금액
            pGDColumn[87] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SMALL_CORPOR_DED_AMT");     // 소기업/소상공인 공제부금 소득공제

            pGDColumn[88] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PER_DED_AMT");              // 기본(본인)
            pGDColumn[89] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_APP_SAVE_AMT");       // 청약저축

            pGDColumn[90] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SPOUSE_DED_AMT");           // 기본(배우자)
            pGDColumn[91] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_APP_DEPOSIT_AMT");    // 주택청약종합저축

            pGDColumn[92] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUPP_DED_COUNT");           // 기본(부양인원 - 인원)            
            pGDColumn[93] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUPP_DED_AMT");             // 기본(부양인원 - 금액)
            pGDColumn[94] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_SAVE_AMT");           // 장기주택마련저축

            pGDColumn[95] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OLD_DED_COUNT");            // 추가공제(경로수 - 인원)          
            pGDColumn[96] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OLD_DED_AMT");              // 추가공제(경로수 - 금액)
            pGDColumn[97] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORKER_HOUSE_SAVE_AMT");    // 근로자주택마련저축

            pGDColumn[98] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_DED_COUNT");         // 추가공제(장애인 - 인원)          
            pGDColumn[99] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_DED_AMT");           // 추가공제(장애인 - 금액)
            pGDColumn[100] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INVES_AMT");               // 투자조합출자등 소득공제

            pGDColumn[101] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WOMAN_DED_AMT");           // 추가공제(부녀세대)
            pGDColumn[102] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CREDIT_AMT");              // 신용카드등 소득공제

            pGDColumn[103] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CHILD_DED_COUNT");         // 추가공제(자녀양육 - 인원)        
            pGDColumn[104] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CHILD_DED_AMT");           // 추가공제(자녀양육 - 금액)
            pGDColumn[105] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPL_STOCK_AMT");          // 우리사주출자

            pGDColumn[106] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BIRTH_DED_COUNT");         // 추가공제(출산입양 - 인원)        
            pGDColumn[107] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BIRTH_DED_AMT");           // 추가공제(출산입양 - 금액)
            pGDColumn[108] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_STOCK_SAVING_AMT");   // 장기주식형저축

            pGDColumn[109] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HIRE_KEEP_EMPLOY_AMT");    // 고용유지중소기업소득공제

            pGDColumn[110] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MANY_CHILD_DED_COUNT");    // 다자녀공제(인원)                 
            pGDColumn[111] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MANY_CHILD_DED_AMT");      // 다자녀공제(금액) 

            pGDColumn[112] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATI_ANNU_AMT");           // 국민연금보험료공제               

            pGDColumn[113] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ETC_DED_SUM");             // 그밖의소득공제 계

            pGDColumn[114] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_STD_AMT");             // 종합과세표준

            pGDColumn[115] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("COMP_TAX_AMT");            // 산출세액

            pGDColumn[116] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_IN_LAW_AMT");     // 소득세법

            pGDColumn[117] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETR_ANNU_AMT");           // 퇴직연금소득공제                 
            pGDColumn[118] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SP_LAW_AMT");     // 조세특례제한법

            pGDColumn[119] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MEDIC_INSUR_AMT");         // 건강보험료                       
            pGDColumn[120] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HIRE_INSUR_AMT");          // 고용보험료                       
            pGDColumn[121] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("GUAR_INSUR_AMT");          // 보장성보험                       
            pGDColumn[122] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_INSUR_AMT");        // 장애인전용 

            pGDColumn[123] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MEDIC_AMT");               // 특별공제(의료비) 
            pGDColumn[124] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SUM");            // 세액감면 계  

            pGDColumn[125] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EDUCATION_AMT");           // 특별공제(교육비) 
            pGDColumn[126] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_INCOME_AMT");      // 근로소득                

            pGDColumn[127] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_INTER_AMT");         // 특별공제(주택임차차입금) 
            pGDColumn[128] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_TAXGROUP_AMT");    // 납세조합공제

            pGDColumn[129] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_MONTHLY_AMT");       // 특별공제(월세)

            pGDColumn[130] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT");     // 특별공제(장기주택차입금-2011)
            pGDColumn[131] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_HOUSE_DEBT_AMT");  // 주택차입금

            pGDColumn[132] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DONAT_AMT");               // 특별공제(기부금) 
            pGDColumn[133] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_DONAT_POLI_AMT");  // 기부 정치자금

            pGDColumn[134] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_OUTSIDE_PAY_AMT"); // 외국 납부

            pGDColumn[135] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SP_DED_SUM");              // 특별공제(계)

            pGDColumn[136] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STAND_DED_AMT");           // 특별공제(표준공제)
            pGDColumn[137] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_SUM");             // 세액공제 계         

            pGDColumn[138] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_DED_AMT");            // 차감소득금액                      
            pGDColumn[139] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SET_TAX_SUM");             // 결정세액                         

            pGDColumn[140] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_3");   // 특별공제(장기주택차입금-2012)

            pGDColumn[141] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECIPIENT_PERSON_NAME");       // 받는자 소득자
            pGDColumn[142] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECIPIENT_TAX_OFFICE_NAME");   // 받는자 세무서
            pGDColumn[143] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECIPIENT_ETC_NAME");          // 받는자 기타 

            //----[ 1 page ]------------------------------------------------------------------------------------------------------
            pXLColumn[0] = 37;  // 거주 구분(거주자1/거주자2)    
            pXLColumn[1] = 37;  // 내외국인 구분(내국인1/외국인9)
            pXLColumn[2] = 39;  // 외국인단일세율적용            
            pXLColumn[3] = 37;  // 세대주 구분(세대주1/세대원2)  
            pXLColumn[4] = 37;  // 연말정산구분                  

            pXLColumn[5] = 11;  // 법인명(상호)                  
            pXLColumn[6] = 30;  // 대표자(성명)                  
            pXLColumn[7] = 11;  // 사업자등록번호                
            pXLColumn[8] = 11;  // 소재지(주소)                  

            pXLColumn[9] = 11;  // 성명                          
            pXLColumn[10] = 30;  // 주민번호                      
            pXLColumn[11] = 11;  // 주소                          

            //--------------------------------------------------------------------------------------------------------------------
            // I 근무처별 소득 명세
            //--------------------------------------------------------------------------------------------------------------------
            pXLColumn[12] = 11;  // 주(현)근무처명                
            pXLColumn[13] = 17;  // 종(전)1근무처명               
            pXLColumn[14] = 23;  // 종(전)2근무처명               

            pXLColumn[15] = 11;  // 주(현)사업자번호              
            pXLColumn[16] = 17;  // 종(전)1사업잡번호             
            pXLColumn[17] = 23;  // 종(전)2사업잡번호             

            pXLColumn[18] = 11;  // 주(현)근무기간                
            pXLColumn[19] = 17;  // 종(전)1근무기간               
            pXLColumn[20] = 23;  // 종(전)2근무기간               

            pXLColumn[21] = 11;  // 주(현)감면기간                
            pXLColumn[22] = 17;  // 종(전)1감면기간               
            pXLColumn[23] = 23;  // 종(전)2감면기간               

            pXLColumn[24] = 11;  // 주(현)급여                    
            pXLColumn[25] = 17;  // 종(전)1급여                   
            pXLColumn[26] = 23;  // 종(전)2급여                   

            pXLColumn[27] = 11;  // 주(현)상여                    
            pXLColumn[28] = 17;  // 종(전)1상여                   
            pXLColumn[29] = 23;  // 종(전)2상여                   

            pXLColumn[30] = 11;  // 주(현)인정상여                
            pXLColumn[31] = 17;  // 종(전)1인정상여               
            pXLColumn[32] = 23;  // 종(전)2인정상여               

            pXLColumn[33] = 11;  // 주(현)주식매수선택권          
            pXLColumn[34] = 17;  // 종(전)1주식매수선택권         
            pXLColumn[35] = 23;  // 종(전)2주식매수선택권         

            pXLColumn[36] = 11;  // 주(현)우리사주조합인출금      
            pXLColumn[37] = 17;  // 종(전)1우리사주조합인출금     
            pXLColumn[38] = 23;  // 종(전)2우리사주조합인출금     

            pXLColumn[39] = 11;  // 주(현)계                      
            pXLColumn[40] = 17;  // 종(전)1계                     
            pXLColumn[41] = 23;  // 종(전)2계                     

            //--------------------------------------------------------------------------------------------------------------------
            // II 비과세 및 감면 소득 명세
            //--------------------------------------------------------------------------------------------------------------------
            pXLColumn[42] = 11;  // 비과세_주(현)국외근로       
            pXLColumn[43] = 17;  // 비과세_종(전)1국외근로      
            pXLColumn[44] = 23;  // 비과세_종(전)2국외근로      

            pXLColumn[45] = 11;  // 비과세_주(현)야간근로수당   
            pXLColumn[46] = 17;  // 비과세_종(전)1야간근로수당  
            pXLColumn[47] = 23;  // 비과세_종(전)2야간근로수당  

            pXLColumn[48] = 11;  // 비과세_주(현)출산/보육수당  
            pXLColumn[49] = 17;  // 비과세_종(전)1출산/보육수당 
            pXLColumn[50] = 23;  // 비과세_종(전)2출산/보육수당 

            pXLColumn[51] = 11;  // 비과세_주(현)외국인근로자   
            pXLColumn[52] = 17;  // 비과세_종(전)1외국인근로자  
            pXLColumn[53] = 23;  // 비과세_종(전)2외국인근로자  

            pXLColumn[54] = 11;  // 비과세_주(현)비과세소득계   
            pXLColumn[55] = 17;  // 비과세_종(전)1비과세소득계  
            pXLColumn[56] = 23;  // 비과세_종(전)2비과세소득계  

            pXLColumn[57] = 11;  // 비과세_주(현)감면소득계     
            pXLColumn[58] = 17;  // 비과세_종(전)1감면소득계    
            pXLColumn[59] = 23;  // 비과세_종(전)2감면소득계    

            //--------------------------------------------------------------------------------------------------------------------
            // III 세액 명세
            //--------------------------------------------------------------------------------------------------------------------
            pXLColumn[60] = 19;  // 결정세액_소득세              
            pXLColumn[61] = 25;  // 결정세액_지방소득세          
            pXLColumn[62] = 31;  // 결정세액_농특세              
            pXLColumn[63] = 36;  // 결정세액_계                  

            pXLColumn[64] = 14;  // 기납부세액_종(전)1사업자번호 
            pXLColumn[65] = 19;  // 기납부세액_종(전)1소득세     
            pXLColumn[66] = 25;  // 기납부세액_종(전)1지방소득세 
            pXLColumn[67] = 31;  // 기납부세액_종(전)1농특세     
            pXLColumn[68] = 36;  // 기납부세액_종(전)1계         

            pXLColumn[69] = 14;  // 기납부세액_종(전)2사업자번호 
            pXLColumn[70] = 19;  // 기납부세액_종(전)2소득세     
            pXLColumn[71] = 25;  // 기납부세액_종(전)2지방소득세 
            pXLColumn[72] = 31;  // 기납부세액_종(전)2농특세                  
            pXLColumn[73] = 36;  // 기납부세액_종(전)2계         

            pXLColumn[74] = 19;  // 기납부세액_주(현)소득세      
            pXLColumn[75] = 25;  // 기납부세액_주(현)지방소득세  
            pXLColumn[76] = 31;  // 기납부세액_주(현)농특세      
            pXLColumn[77] = 36;  // 기납부세액_주(현)계          

            pXLColumn[78] = 19;  // 차감징수세액_소득세          
            pXLColumn[79] = 25;  // 차감징수세액_지방소득세      
            pXLColumn[80] = 31;  // 차감징수세액_농특세          
            pXLColumn[81] = 36;  // 차감징수세액_계  

            //----[ 2 page ]------------------------------------------------------------------------------------------------------
            pXLColumn[82] = 16; // 총급여
            pXLColumn[83] = 36; // 개인연금저축소득공제

            pXLColumn[84] = 16; // 근로소득공제
            pXLColumn[85] = 36; // 연금저축소득공제

            pXLColumn[86] = 16; // 근로소득금액
            pXLColumn[87] = 36; // 소기업/소상공인 공제부금 소득공제

            pXLColumn[88] = 16; // 기본(본인)
            pXLColumn[89] = 36; // 청약저축

            pXLColumn[90] = 16; // 기본(배우자)
            pXLColumn[91] = 36; // 주택청약종합저축

            pXLColumn[92] = 10; // 기본(부양인원 - 인원)            
            pXLColumn[93] = 16; // 기본(부양인원 - 금액)
            pXLColumn[94] = 36; // 장기주택마련저축

            pXLColumn[95] = 10; // 추가공제(경로수 - 인원)          
            pXLColumn[96] = 16; // 추가공제(경로수 - 금액)
            pXLColumn[97] = 36; // 근로자주택마련저축

            pXLColumn[98] = 10; // 추가공제(장애인 - 인원)          
            pXLColumn[99] = 16; // 추가공제(장애인 - 금액)
            pXLColumn[100] = 36; // 투자조합출자등 소득공제

            pXLColumn[101] = 16; // 추가공제(부녀세대)
            pXLColumn[102] = 36; // 신용카드등 소득공제

            pXLColumn[103] = 10; // 추가공제(자녀양육 - 인원)        
            pXLColumn[104] = 16; // 추가공제(자녀양육 - 금액)
            pXLColumn[105] = 36; // 우리사주출자

            pXLColumn[106] = 11; // 추가공제(출산입양 - 인원)        
            pXLColumn[107] = 16; // 추가공제(출산입양 - 금액)
            pXLColumn[108] = 36; // 장기주식형저축

            pXLColumn[109] = 36; // 고용유지중소기업소득공제

            pXLColumn[110] = 9; // 다자녀공제(인원)                 
            pXLColumn[111] = 16; // 다자녀공제(금액) 

            pXLColumn[112] = 16; // 국민연금보험료공제               

            pXLColumn[113] = 36; // 그밖의소득공제 계

            pXLColumn[114] = 36; // 종합과세표준

            pXLColumn[115] = 36; // 산출세액

            pXLColumn[116] = 36; // 소득세법

            pXLColumn[117] = 16; // 퇴직연금소득공제                 
            pXLColumn[118] = 36; // 조세특례제한법

            pXLColumn[119] = 16; // 건강보험료                       
            pXLColumn[120] = 16; // 고용보험료                       
            pXLColumn[121] = 16; // 보장성보험                       
            pXLColumn[122] = 16; // 장애인전용 

            pXLColumn[123] = 16; // 특별공제(의료비) 
            pXLColumn[124] = 36; // 세액감면 계  

            pXLColumn[125] = 16; // 특별공제(교육비) 
            pXLColumn[126] = 36; // 근로소득                

            pXLColumn[127] = 16; // 특별공제(주택임차차입금) 
            pXLColumn[128] = 36; // 납세조합공제

            pXLColumn[129] = 16; // 특별공제(월세)

            pXLColumn[130] = 16; // 특별공제(장기주택차입금)
            pXLColumn[131] = 36; // 주택차입금

            pXLColumn[132] = 16; // 특별공제(기부금) 
            pXLColumn[133] = 36; // 기부 정치자금

            pXLColumn[134] = 36; // 외국 납부

            pXLColumn[135] = 16; // 계

            pXLColumn[136] = 16; // 특별공제(표준공제)
            pXLColumn[137] = 36; // 세액공제 계         

            pXLColumn[138] = 16; // 차감소득금액  
            pXLColumn[139] = 36; // 결정세액    

            pXLColumn[140] = 16; // 결정세액   

            pXLColumn[141] = 2; // 소득자 보관   
            pXLColumn[142] = 2; // 세무서 제출
            pXLColumn[143] = 2; // 발행자 보관 
        }

        #endregion;

        #region ----- Array Set 2 ----
        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_SUPPORT_FAMILY, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[60];
            pXLColumn = new int[60];

            pGDColumn[0] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("MANY_CHILD_DED_COUNT");  // 다자녀 인원수
            pGDColumn[1] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("RELATION_CODE");         // 관계코드         
            pGDColumn[2] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("FAMILY_NAME");           // 성명       

            pGDColumn[3] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BASE_YN");               // 기본공제         
            pGDColumn[4] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("OLD_YN");                // 경로우대         
            pGDColumn[5] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BIRTH_YN");              // 출산/입양양육    
            pGDColumn[6] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DISABILITY_YN");         // 장애인           
            pGDColumn[7] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHILD_YN");              // 자녀양육(6세이하)
            pGDColumn[8] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("INSURE_AMT");            // 국세청-보험료    
            pGDColumn[9] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("MEDICAL_AMT");           // 국세청-의료비    
            pGDColumn[10] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("EDU_AMT");               // 국세청-교육비    
            pGDColumn[11] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CREDIT_AMT");            // 국세청-신용카드  
            pGDColumn[12] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHECK_CREDIT_AMT");      // 국세청-직불카드  
            pGDColumn[13] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CASH_AMT");              // 국세청-현금      
            pGDColumn[14] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DONAT_AMT");             // 국세청-기부금    
            pGDColumn[15] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("NATIONALITY_TYPE");      // 국가타입         
            pGDColumn[16] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("REPRE_NUM");             // 주민번호         
            pGDColumn[17] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("WOMAN_YN");              // 부녀자           
            pGDColumn[18] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_INSURE_AMT");        // 기타-보험료      
            pGDColumn[19] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_MEDICAL_AMT");       // 기타-의료비      
            pGDColumn[20] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_EDU_AMT");           // 기타-교육비      
            pGDColumn[21] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_CREDIT_AMT");        // 기타-신용카드    
            pGDColumn[22] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHECK_ETC_CREDIT_AMT");  // 기타-직불카드    
            pGDColumn[23] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_CASH_AMT");          // 기타-현금        
            pGDColumn[24] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_DONAT_AMT");         // 기타-기부금 

            pGDColumn[25] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BASE_COUNT");            // 기본공제
            pGDColumn[26] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("OLD_COUNT");             // 경로우대
            pGDColumn[27] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BIRTH_COUNT");           // 출산/입양양육
            pGDColumn[28] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DISABILITY_COUNT");      // 장애인
            pGDColumn[29] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHILD_COUNT");           // 자녀양육
            pGDColumn[30] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("WOMAN_COUNT");           // 부녀세대

            //2013추가//
            pGDColumn[31] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BASE_YN");              // 기본공제
            pGDColumn[32] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("OLD_YN");               // 경로우대
            pGDColumn[33] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BIRTH_YN");             // 출산/입양양육
            pGDColumn[34] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("WOMAN_YN");             // 부녀자
            pGDColumn[35] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DISABILITY_YN");        // 장애인
            pGDColumn[36] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHILD_YN");             // 6세이하

            //2013변경//
            //pGDColumn[1] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("RELATION_CODE");         // 관계코드         
            //pGDColumn[2] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("FAMILY_NAME");           // 성명      
            pGDColumn[37] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("INSURE_AMT");           // 국세청-보험료
            pGDColumn[38] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("MEDICAL_AMT");          // 국세청-의료비
            pGDColumn[39] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("EDU_AMT");              // 국세청-교육비
            pGDColumn[40] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CREDIT_AMT");           // 국세청-신용카드
            pGDColumn[41] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHECK_CREDIT_AMT");     // 국세청-직불카드
            pGDColumn[42] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ACADE_GIRO_AMT");       // 국세청-학원비지로납부액
            pGDColumn[43] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CASH_AMT");             // 국세청-현금영수증
            pGDColumn[44] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("TRAD_MARKET_AMT");      // 국세청-전통시장사용액
            pGDColumn[45] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DONAT_AMT");            // 국세청-기부금

            pGDColumn[46] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("NATIONALITY_TYPE");      // 국가타입
            pGDColumn[47] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("REPRE_NUM");             // 주민번호

            pGDColumn[48] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_INSURE_AMT");        // 기타-보험료
            pGDColumn[49] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_MEDICAL_AMT");       // 기타-의료비
            pGDColumn[50] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_EDU_AMT");           // 기타-교육비
            pGDColumn[51] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_CREDIT_AMT");        // 기타-신용카드
            pGDColumn[52] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHECK_ETC_CREDIT_AMT");  // 기타-직불카드
            pGDColumn[53] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_ACADE_GIRO_AMT");    // 기타-학원비지로납부액
            pGDColumn[54] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_CASH_AMT");          // 기타-현금영수증
            pGDColumn[55] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_TRAD_MARKET_AMT");   // 기타-전통시장사용액
            pGDColumn[56] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_DONAT_AMT");         // 기타-기부금



            //---------------------------------------------------------------------------------------------------------------
            pXLColumn[0] = 7;   // 다자녀 인원수
            pXLColumn[1] = 1;   // 관계코드         
            pXLColumn[2] = 3;   // 성명     

            pXLColumn[3] = 9;   // 기본공제         
            pXLColumn[4] = 11;  // 경로우대         
            pXLColumn[5] = 13;  // 출산/입양양육    
            pXLColumn[6] = 15;  // 장애인           
            pXLColumn[7] = 17;  // 자녀양육(6세이하)
            pXLColumn[8] = 22;  // 국세청-보험료    
            pXLColumn[9] = 25;  // 국세청-의료비    
            pXLColumn[10] = 28;  // 국세청-교육비    
            pXLColumn[11] = 31;  // 국세청-신용카드  
            pXLColumn[12] = 34;  // 국세청-직불카드  
            pXLColumn[13] = 37;  // 국세청-현금      
            pXLColumn[14] = 40;  // 국세청-기부금    
            pXLColumn[15] = 1;   // 국가타입         
            pXLColumn[16] = 3;   // 주민번호         
            pXLColumn[17] = 9;   // 부녀자           
            pXLColumn[18] = 22;  // 기타-보험료      
            pXLColumn[19] = 25;  // 기타-의료비      
            pXLColumn[20] = 28;  // 기타-교육비      
            pXLColumn[21] = 31;  // 기타-신용카드    
            pXLColumn[22] = 34;  // 기타-직불카드    
            pXLColumn[23] = 37;  // 기타-현금        
            pXLColumn[24] = 40;  // 기타-기부금
            pXLColumn[25] = 9;   // 기본공제         
            pXLColumn[26] = 11;  // 경로우대         
            pXLColumn[27] = 13;  // 출산/입양양육    
            pXLColumn[28] = 15;  // 장애인           
            pXLColumn[29] = 17;  // 자녀양육(6세이하)
            pXLColumn[30] = 9;   // 부녀세대

            //2013년추가//
            pXLColumn[31] = 9;   // 기본공제         
            pXLColumn[32] = 11;  // 경로우대         
            pXLColumn[33] = 13;  // 출산/입양양육  
            pXLColumn[34] = 9;   // 부녀자
            pXLColumn[35] = 11;  // 장애인           
            pXLColumn[36] = 13;  // 6세이하

            //2013년변경//
            //pXLColumn[0] = 7;   // 다자녀 인원수
            //pXLColumn[1] = 1;   // 관계코드         
            //pXLColumn[2] = 3;   // 성명    
            pXLColumn[37] = 17;  // 국세청-보험료    
            pXLColumn[38] = 20;  // 국세청-의료비    
            pXLColumn[39] = 23;  // 국세청-교육비    
            pXLColumn[40] = 26;  // 국세청-신용카드  
            pXLColumn[41] = 29;  // 국세청-직불카드  
            pXLColumn[42] = 32;  // 국세청-학원비지로납부액      
            pXLColumn[43] = 35;  // 국세청-현금영수증
            pXLColumn[44] = 38;  // 국세청-전통시장사용액 
            pXLColumn[45] = 41;  // 국세청-기부금
            pXLColumn[46] = 1;   // 국가타입  
            pXLColumn[47] = 3;   // 주민번호     
            pXLColumn[48] = 17;  // 기타-보험료      
            pXLColumn[49] = 20;  // 기타-의료비      
            pXLColumn[50] = 23;  // 기타-교육비      
            pXLColumn[51] = 26;  // 기타-신용카드    
            pXLColumn[52] = 29;  // 기타-직불카드    
            pXLColumn[53] = 32;  // 기타-학원비지로납부액    
            pXLColumn[54] = 35;  // 기타-현금영수증
            pXLColumn[55] = 38;  // 기타-전통시장사용액  
            pXLColumn[56] = 41;  // 기타-전통시장사용액  

        }

        #endregion;

        #region -----  Array Set 3 -----

        private void SetArray3(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[5];
            pXLColumn = new int[5];

            pGDColumn[0] = pGrid.GetColumnToIndex("SAVING_TYPE_NAME");
            pGDColumn[1] = pGrid.GetColumnToIndex("BANK_NAME");
            pGDColumn[2] = pGrid.GetColumnToIndex("ACCOUNT_NUM");
            pGDColumn[3] = pGrid.GetColumnToIndex("SAVING_AMOUNT");
            pGDColumn[4] = pGrid.GetColumnToIndex("SAVING_DED_AMOUNT");


            pXLColumn[0] = 1;   //SAVING_TYPE_NAME
            pXLColumn[1] = 8;   //BANK_NAME
            pXLColumn[2] = 17;  //ACCOUNT_NUM
            pXLColumn[3] = 26;  //SAVING_AMOUNT
            pXLColumn[4] = 35;  //SAVING_DED_AMOUNT
        }

        //private void SetArray3(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        //{
        //    pGDColumn = new int[5];
        //    pXLColumn = new int[5];

        //    pGDColumn[0] = pTable.Columns.IndexOf("SAVING_TYPE_NAME");
        //    pGDColumn[1] = pTable.Columns.IndexOf("BANK_NAME");
        //    pGDColumn[2] = pTable.Columns.IndexOf("ACCOUNT_NUM");
        //    pGDColumn[3] = pTable.Columns.IndexOf("SAVING_AMOUNT");
        //    pGDColumn[4] = pTable.Columns.IndexOf("SAVING_DED_AMOUNT");


        //    pXLColumn[0] = 1;   //SAVING_TYPE_NAME
        //    pXLColumn[1] = 8;   //BANK_NAME
        //    pXLColumn[2] = 17;  //ACCOUNT_NUM
        //    pXLColumn[3] = 26;  //SAVING_AMOUNT
        //    pXLColumn[4] = 35;  //SAVING_DED_AMOUNT
        //}

        #endregion;

        #region ----- Array Set 4 -----

        private void SetArray4(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[5];
            pXLColumn = new int[5];

            pGDColumn[0] = pGrid.GetColumnToIndex("BANK_NAME");
            pGDColumn[1] = pGrid.GetColumnToIndex("ACCOUNT_NUM");
            pGDColumn[2] = pGrid.GetColumnToIndex("SAVING_COUNT");
            pGDColumn[3] = pGrid.GetColumnToIndex("SAVING_AMOUNT");
            pGDColumn[4] = pGrid.GetColumnToIndex("SAVING_DED_AMOUNT");


            pXLColumn[0] = 1;   //BANK_NAME
            pXLColumn[1] = 8;   //ACCOUNT_NUM
            pXLColumn[2] = 17;  //SAVING_COUNT
            pXLColumn[3] = 26;  //SAVING_AMOUNT
            pXLColumn[4] = 35;  //SAVING_DED_AMOUNT
        }

        #endregion;

        #region ----- Array Set 5 ----

        private void SetArray5(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[191];
            pXLColumn = new int[191];

            //-----------------------------------------------------------------------------------------------------------------------------------
            //-- 1. 오른쪽 상단 표 
            //-----------------------------------------------------------------------------------------------------------------------------------
            pGDColumn[0] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TYPE");        // 거주 구분(거주자1/거주자2)  

            pGDColumn[1] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_NAME");        // 거주지국
            pGDColumn[2] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_CODE");        // 거주지국코드

            pGDColumn[3] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONALITY_TYPE");     // 내외국인 구분(내국인1/외국인9)

            pGDColumn[4] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FOREIGN_TAX_YN");       // 외국인단일세율적용

            pGDColumn[5] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONAL_NAME");        // 국적
            pGDColumn[6] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONAL_CODE");        // 국적코드

            pGDColumn[7] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSEHOLD_TYPE");       // 세대주 구분(세대주1/세대원2)    

            pGDColumn[8] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_KEEP_TYPE");       // 연말정산구분


            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 2. 징수 의무자
            //-----------------------------------------------------------------------------------------------------------------------------------  
            pGDColumn[9] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CORP_NAME");            // 법인명(상호)                  
            pGDColumn[10] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRESIDENT_NAME");      // 대표자(성명)      

            pGDColumn[11] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("VAT_NUMBER");          // 사업자등록번호   

            pGDColumn[12] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ORG_ADDRESS");         // 소재지(주소)

            //-----------------------------------------------------------------------------------------------------------------------------------
            // --3.소득자
            // ---------------------------------------------------------------------------------------------------------------------------------- -  
            pGDColumn[13] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NAME");                // 성명
            pGDColumn[14] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPRE_NUM");           // 주민번호

            pGDColumn[15] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PERSON_ADDRESS");      // 주소  

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 4. 근무처별소득명세 : 
            // -----------------------------------------------------------------------------------------------------------------------------------        
            pGDColumn[16] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME");      // 주(현)근무처명
            pGDColumn[17] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NAME1");    // 종(전)1근무처명 
            pGDColumn[18] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NAME2");    // 종(전)2근무처명

            pGDColumn[19] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER");     // 주(현)사업자번호 
            pGDColumn[20] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NUM1");     // 종(전)1사업자번호
            pGDColumn[21] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NUM2");     // 종(전)2사업자번호 

            pGDColumn[22] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE");         // 주(현)근무기간
            pGDColumn[23] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE1");        // 종(전)1근무기간
            pGDColumn[24] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE2");        // 종(전)2근무기간

            pGDColumn[25] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE");         // 주(현)감면기간 
            pGDColumn[26] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE1");        // 종(전)1감면기간
            pGDColumn[27] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE2");        // 종(전)2감면기간

            pGDColumn[28] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_PAY_TOT_AMT");     // 주(현)급여 
            pGDColumn[29] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PAY_TOTAL_AMT1");      // 종(전)1급여
            pGDColumn[30] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PAY_TOTAL_AMT2");      // 종(전)2급여

            pGDColumn[31] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_BONUS_TOT_AMT");   // 주(현)상여   
            pGDColumn[32] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BONUS_TOTAL_AMT1");    // 종(전)1상여 
            pGDColumn[33] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BONUS_TOTAL_AMT2");    // 종(전)2상여

            pGDColumn[34] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_ADD_BONUS_AMT");   // 주(현)인정상여   
            pGDColumn[35] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_BONUS_AMT1");      // 종(전)1인정상여
            pGDColumn[36] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_BONUS_AMT2");      // 종(전)2인정상여

            pGDColumn[37] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_STOCK_BENE_AMT");  // 주(현)주식매수선택권
            pGDColumn[38] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STOCK_BENE_AMT1");     // 종(전)1주식매수선택권
            pGDColumn[39] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STOCK_BENE_AMT2");     // 종(전)2주식매수선택권

            pGDColumn[40] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_EMPLOYEE_STOCK_AMT");       // 주(현)우리사주조합인출금
            pGDColumn[41] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPLOYEE_STOCK_AMT1");       // 종(전)1우리사주조합인출금
            pGDColumn[42] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPLOYEE_STOCK_AMT2");               // 종(전)2우리사주조합인출금

            pGDColumn[43] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_OFFICE_RETIRE_OVER_AMT");    // 주(현)임원퇴직소득금액 한도초과액
            pGDColumn[44] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OFFICE_RETIRE_OVER_AMT1");   // 종(전)1임원퇴직소득금액 한도초과액
            pGDColumn[45] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OFFICE_RETIRE_OVER_AMT2");   // 종(전)2임원퇴직소득금액 한도초과액

            pGDColumn[46] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_TOTAL_AMOUNT");    // 주(현)계     
            pGDColumn[47] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AMOUNT1");       // 종(전)1계   
            pGDColumn[48] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AMOUNT2");       // 종(전)2계  

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 5. 비과세 및 감면소득 명세
            // -----------------------------------------------------------------------------------------------------------------------------------
            pGDColumn[49] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_OUTSIDE_AMT");  // 비과세_주(현)국외근로
            pGDColumn[50] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OUTSIDE_AMT1");     // 비과세_종(전)1국외근로
            pGDColumn[51] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OUTSIDE_AMT2");     // 비과세_종(전)2국외근로

            pGDColumn[52] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_OT_AMT");       // 비과세_주(현)야간근로수당 
            pGDColumn[53] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OT_AMT1");          // 비과세_종(전)1야간근로수당
            pGDColumn[54] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OT_AMT2");          // 비과세_종(전)2야간근로수당

            pGDColumn[55] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_BIRTH_AMT");    // 비과세_주(현)출산/보육수당
            pGDColumn[56] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_BIRTH_AMT1");       // 비과세_종(전)1출산/보육수당
            pGDColumn[57] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_BIRTH_AMT2");       // 비과세_종(전)2출산/보육수당

            pGDColumn[58] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_COMPANY_AMT");   // 비과세_주(현)연구보조비
            pGDColumn[59] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_COMPANY_AMT1");  // 비과세_종(전)1연구보조비
            pGDColumn[60] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_COMPANY_AMT2");  // 비과세_종(전)2연구보조비

            pGDColumn[61] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TRAIN_AMT");     // 비과세_주(현)수련보조수당
            pGDColumn[62] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TRAIN_AMT1");    // 비과세_종(전)1수련보조수당
            pGDColumn[63] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TRAIN_AMT2");    // 비과세_종(전)2수련보조수당

            pGDColumn[64] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TOTAL_AMOUNT");  // 비과세_주(현)비과세소득 계
            pGDColumn[65] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_TOTAL_AMOUNT1");     // 비과세_종(전)1비과세소득 계
            pGDColumn[66] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_TOTAL_AMOUNT2");     // 비과세_종(전)2비과세소득 계

            pGDColumn[67] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT");  // 비과세_주(현)감면소득 계
            pGDColumn[68] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT1"); // 비과세_종(전)1감면소득 계
            pGDColumn[69] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT2"); // 비과세_종(전)2감면소득 계

            //--------------------------------------------------------------------------------------------------------------------
            // 6. 세액 명세
            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[70] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_IN_TAX_AMT");      // 결정세액_소득세               
            pGDColumn[71] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_LOCAL_TAX_AMT");   // 결정세액_지방소득세               
            pGDColumn[72] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_SP_TAX_AMT");      // 결정세액_농특세               

            pGDColumn[73] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_COMPANY_NUM1");    // 기납부세액_종(전)1사업자번호  
            pGDColumn[74] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_IN_TAX_AMT1");     // 기납부세액_종(전)1소득세      
            pGDColumn[75] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_LOCAL_TAX_AMT1");  // 기납부세액_종(전)1지방소득세      
            pGDColumn[76] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_SP_TAX_AMT1");     // 기납부세액_종(전)1농특세      


            pGDColumn[77] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_COMPANY_NUM2");    // 기납부세액_종(전)2사업자번호  
            pGDColumn[78] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_IN_TAX_AMT2");     // 기납부세액_종(전)2소득세      
            pGDColumn[79] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_LOCAL_TAX_AMT2");  // 기납부세액_종(전)2지방소득세      
            pGDColumn[80] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_SP_TAX_AMT2");     // 기납부세액_종(전)2농특세      

            pGDColumn[81] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_IN_TAX_AMT");      // 기납부세액_주(현)소득세       
            pGDColumn[82] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LOCAL_TAX_AMT");   // 기납부세액_주(현)지방소득세       
            pGDColumn[83] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_SP_TAX_AMT");      // 기납부세액_주(현)농특세       

            pGDColumn[84] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT");     // 차감징수세액_소득세           
            pGDColumn[85] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT");  // 차감징수세액_지방소득세         
            pGDColumn[86] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_SP_TAX_AMT");     // 차감징수세액_농특세           


            //--------------------------------------------------------------------------------------------------------------------
            //[ 2 page ]
            //--------------------------------------------------------------------------------------------------------------------

            pGDColumn[87] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_TOT_AMT");           // 총급여
            pGDColumn[88] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_DED_AMT");           // 근로소득공제
            pGDColumn[89] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_AMT");               // 근로소득금액

            // 기본공제
            pGDColumn[90] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PER_DED_AMT");              // 기본(본인)
            pGDColumn[91] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SPOUSE_DED_AMT");           // 기본(배우자)
            pGDColumn[92] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUPP_DED_COUNT");           // 기본(부양인원 - 인원)            
            pGDColumn[93] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUPP_DED_AMT");             // 기본(부양인원 - 금액)

            // 추가공제
            pGDColumn[94] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OLD_DED_COUNT");            // 추가공제(경로수 - 인원)          
            pGDColumn[95] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OLD_DED_AMT");              // 추가공제(경로수 - 금액)
            pGDColumn[96] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_DED_COUNT");     // 추가공제(장애인 - 인원)          
            pGDColumn[97] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_DED_AMT");       // 추가공제(장애인 - 금액)
            pGDColumn[98] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WOMAN_DED_AMT");            // 추가공제(부녀세대)
            pGDColumn[99] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CHILD_DED_COUNT");          // 추가공제(자녀양육 - 인원)        
            pGDColumn[100] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CHILD_DED_AMT");            // 추가공제(자녀양육 - 금액)
            pGDColumn[101] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BIRTH_DED_COUNT");          // 추가공제(출산입양 - 인원)        
            pGDColumn[102] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BIRTH_DED_AMT");           // 추가공제(출산입양 - 금액)
            pGDColumn[103] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SINGLE_PARENT_DED_AMT");   // 추가공제(한부모가족)

            pGDColumn[104] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MANY_CHILD_DED_COUNT");    // 다자녀공제(인원)                 
            pGDColumn[105] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MANY_CHILD_DED_AMT");      // 다자녀공제(금액) 

            // 연금보험료공제 
            pGDColumn[106] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATI_ANNU_AMT");           // 국민연금보험료공제               

            pGDColumn[107] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PUBLIC_INSUR_AMT");        // 공무원 연금  
            pGDColumn[108] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MARINE_INSUR_AMT");        // 군인연금
            pGDColumn[109] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SCHOOL_STAFF_INSUR_AMT");  // 사립학교 교직원 연금
            pGDColumn[110] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("POST_OFFICE_INSUR_AMT");   // 별정우체국 연금

            pGDColumn[111] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SCIENTIST_ANNU_AMT");      // 과학기술인공제
            pGDColumn[112] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETR_ANNU_AMT");           // 근로자퇴직급여 보장법에 따른 퇴직연금
            pGDColumn[113] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ANNU_BANK_AMT");           // 연금저축

            //특별소득공제
            pGDColumn[114] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MEDIC_INSUR_AMT");         // 건강보험료                       
            pGDColumn[115] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HIRE_INSUR_AMT");          // 고용보험료                       
            pGDColumn[116] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("GUAR_INSUR_AMT");          // 보장성보험                       
            pGDColumn[117] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_INSUR_AMT");    // 장애인전용 

            pGDColumn[118] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_MEDIC_AMT");    // 의료비 (장애인)              
            pGDColumn[119] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ETC_MEDIC_AMT");           // 의료비 (기타)

            pGDColumn[120] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_EDUCATION_AMT");// 교육비 (장애인)                          
            pGDColumn[121] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ETC_EDUCATION_AMT");       // 교육비 (기타)

            pGDColumn[122] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_INTER_AMT");         // 주택임차차입금원리금상환액 (대출기관) 
            pGDColumn[123] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_INTER_AMT_ETC");     // 주택임차차입금원리금상환액 (거주자)

            pGDColumn[124] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_MONTHLY_AMT");       // 월세액

            pGDColumn[125] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT");     // 장기주택저당차입금이자상환액 - 2011이전 (15년미만)
            pGDColumn[126] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_1");   // 장기주택저당차입금이자상환액 - 2011이전 (15년~29년)
            pGDColumn[127] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_2");   // 장기주택저당차입금이자상환액 - 2011이전 (30년 이상)

            pGDColumn[128] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_3_FIX"); // 장기주택저당차입금이자상환액 - 2012이후 (고정금리비거치 상환대출)
            pGDColumn[129] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_3_ETC"); // 장기주택저당차입금이자상환액 - 2012이후 (기타)

            pGDColumn[130] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_DONAT_POLI_AMT");  // 정치자금기부금                       
            pGDColumn[131] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DONAT_DED_ALL");           // 법정기부금                       
            pGDColumn[132] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DONAT_DED_30");            // 우리사주조합기부금                       
            pGDColumn[133] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DONAT_DED");               // 지정기부금 

            pGDColumn[134] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SP_DED_SUM");              // 계

            pGDColumn[135] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STAND_DED_AMT");           // 표준공제

            //차감소득금액
            pGDColumn[136] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_DED_AMT");            // 차감소득금액


            //그밖의소득공제

            pGDColumn[137] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PERS_ANNU_BANK_AMT");      //개인연금저축소득공제 

            pGDColumn[138] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SMALL_CORPOR_DED_AMT");    // 소기업/소상공인 공제부금 소득공제

            pGDColumn[139] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_APP_SAVE_AMT");      // 청약저축
            pGDColumn[140] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_APP_DEPOSIT_AMT");   // 주택청약종합저축
            pGDColumn[141] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORKER_HOUSE_SAVE_AMT");   // 근로자주택마련저축

            pGDColumn[142] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INVES_AMT");               // 투자조합출자등 소득공제
            pGDColumn[143] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CREDIT_AMT");              // 신용카드등소득공제
            pGDColumn[144] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPL_STOCK_AMT");          // 우리사주조합소득공제
            pGDColumn[145] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HIRE_KEEP_EMPLOY_AMT");    // 우리사주조합소득공제
            pGDColumn[146] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_LEASE_DED_AMT");       // 고용유지중소기업소득공제

            pGDColumn[147] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ETC_DED_SUM");             // 그 밖의 소득공제 계   

            pGDColumn[148] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SP_DED_TOT_AMT");          // 특별공제 종합한도 초과액 

            pGDColumn[149] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_STD_AMT");             // 종합과세표준

            pGDColumn[150] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("COMP_TAX_AMT");            // 산출세액

            //세액감면

            pGDColumn[151] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_IN_LAW_AMT");     // 소득세법
            pGDColumn[152] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SP_LAW_AMT");     // 조세특례제한법 <53>-1 제외 
            pGDColumn[153] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SP_LAW_AMT2");    // 조세특례제한법 제30조 
            pGDColumn[154] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_LAW_AMT");        // 조세조약

            pGDColumn[155] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SUM");            // 세액감면 계
            pGDColumn[156] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_INCOME_AMT");      // 근로소득
            pGDColumn[157] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_TAXGROUP_AMT");    // 납세조합공제
            pGDColumn[158] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_HOUSE_DEBT_AMT");  // 주택차입금
            pGDColumn[159] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_DONAT_POLI_AMT2"); // 기부 정치자금
            pGDColumn[160] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SUM");            // 외국 납부

            pGDColumn[161] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_SUM");             // 세액공제 계     

            //결정세액
            pGDColumn[162] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESULT_TAX_SUM");          // 결정세액

            //그밖의것들
            pGDColumn[163] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WITHHOLDING_OWNER");          // 대표이사


            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 추가 종(전)3 
            // ----------------------------------------------------------------------------------------------------------------------------------- 

            pGDColumn[164] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_EXIST3");         // 종(전)
            pGDColumn[165] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NAME3");          // 종(전)3근무처명
            pGDColumn[166] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NUM3");           // 종(전)3사업자번호 
            pGDColumn[167] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE3");              // 종(전)3근무기간
            pGDColumn[168] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE3");              // 종(전)3감면기간
            pGDColumn[169] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PAY_TOTAL_AMT3");            // 종(전)3급여
            pGDColumn[170] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BONUS_TOTAL_AMT3");          // 종(전)3상여
            pGDColumn[171] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_BONUS_AMT3");            // 종(전)3인정상여
            pGDColumn[172] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STOCK_BENE_AMT3");           // 종(전)3주식매수선택권
            pGDColumn[173] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPLOYEE_STOCK_AMT3");       // 종(전)3우리사주조합인출금
            pGDColumn[174] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OFFICE_RETIRE_OVER_AMT3");   // 종(전)3임원퇴직소득금액 한도초과액
            pGDColumn[175] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AMOUNT3");             // 종(전)3계 


            pGDColumn[176] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OUTSIDE_AMT3");          // 비과세_종(전)3국외근로
            pGDColumn[177] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OT_AMT3");               // 비과세_종(전)3야간근로수당
            pGDColumn[178] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_BIRTH_AMT3");            // 비과세_종(전)3출산/보육수당
            pGDColumn[179] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_COMPANY_AMT3");      // 비과세_종(전)3연구보조비
            pGDColumn[180] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TRAIN_AMT3");        // 비과세_종(전)3수련보조수당
            pGDColumn[181] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_TOTAL_AMOUNT3");         // 비과세_종(전)3출산/보육수당
            pGDColumn[182] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT3");     // 비과세_종(전)3감면소득 계

            pGDColumn[183] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW3_COMPANY_NUM3");        // 기납부세액_종(전)3사업자번호  
            pGDColumn[184] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW3_IN_TAX_AMT3");         // 기납부세액_종(전)3소득세      
            pGDColumn[185] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW3_LOCAL_TAX_AMT3");      // 기납부세액_종(전)3지방소득세      
            pGDColumn[186] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW3_SP_TAX_AMT3");         // 기납부세액_종(전)3농특세   

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 추가 종(전)중소기업에 취업하는 청년에 대한 소득세 감면
            // -----------------------------------------------------------------------------------------------------------------------------------

            pGDColumn[187] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RD_SMALL_BUSINESS_AMT_EXIST");

            pGDColumn[188] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RD_SMALL_BUSINESS_AMT1");
            pGDColumn[189] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RD_SMALL_BUSINESS_AMT2");
            pGDColumn[190] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RD_SMALL_BUSINESS_AMT3");

            //-----------------------------------------------------------------------------------------------------------------------------------
            //-- 1. 오른쪽 상단 표 
            //-----------------------------------------------------------------------------------------------------------------------------------
            pXLColumn[0] = 37;  // 거주 구분(거주자1/거주자2)  

            pXLColumn[1] = 31;  // 거주지국
            pXLColumn[2] = 39;  // 거주지국코드    

            pXLColumn[3] = 37;  // 내외국인 구분(내국인1/외국인9)

            pXLColumn[4] = 39;  // 외국인단일세율적용          

            pXLColumn[5] = 31;  // 국적                
            pXLColumn[6] = 39;  // 국적코드      

            pXLColumn[7] = 37;  // 세대주 구분(세대주1/세대원2)    

            pXLColumn[8] = 37;  // 연말정산구분           

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 2. 징수 의무자
            //-----------------------------------------------------------------------------------------------------------------------------------  
            pXLColumn[9] = 11;   // 법인명(상호)           
            pXLColumn[10] = 30;  // 대표자(성명)  

            pXLColumn[11] = 11;  // 사업자등록번호     

            pXLColumn[12] = 11;  // 소재지(주소)  

            //-----------------------------------------------------------------------------------------------------------------------------------
            // --3.소득자
            // ---------------------------------------------------------------------------------------------------------------------------------- -  
            pXLColumn[13] = 11;  // 성명
            pXLColumn[14] = 30;  // 주민번호

            pXLColumn[15] = 11;  // 주소 

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 4. 근무처별소득명세 : 
            // -----------------------------------------------------------------------------------------------------------------------------------  
            pXLColumn[16] = 11;  // 주(현)근무처명 
            pXLColumn[17] = 17;  // 종(전)1근무처명 
            pXLColumn[18] = 23;  // 종(전)2근무처명

            pXLColumn[19] = 11;  // 주(현)사업자번호 
            pXLColumn[20] = 17;  // 종(전)1사업자번호 
            pXLColumn[21] = 23;  // 종(전)2사업자번호 

            pXLColumn[22] = 11;  // 주(현)근무기간 
            pXLColumn[23] = 17;  // 종(전)1근무기간 
            pXLColumn[24] = 23;  // 종(전)2근무기간 

            pXLColumn[25] = 11;  // 주(현)감면기간 
            pXLColumn[26] = 17;  // 종(전)1감면기간 
            pXLColumn[27] = 23;  // 종(전)2감면기간 

            pXLColumn[28] = 11;  // 주(현)급여 
            pXLColumn[29] = 17;  // 종(전)1급여 
            pXLColumn[30] = 23;  // 종(전)2급여 

            pXLColumn[31] = 11;  // 주(현)상여 
            pXLColumn[32] = 17;  // 종(전)1상여 
            pXLColumn[33] = 23;  // 종(전)2상여 

            pXLColumn[34] = 11;  // 주(현) 인정상여 
            pXLColumn[35] = 17;  // 종(전)1인정상여 
            pXLColumn[36] = 23;  // 종(전)2인정상여 

            pXLColumn[37] = 11;  // 주(현) 주식매수선택권 
            pXLColumn[38] = 17;  // 종(전)1주식매수선택권 
            pXLColumn[39] = 23;  // 종(전)2주식매수선택권 

            pXLColumn[40] = 11;  // 주(현) 우리사주조합인출금 
            pXLColumn[41] = 17;  // 종(전)1우리사주조합인출금 
            pXLColumn[42] = 23;  // 종(전)2우리사주조합인출금 

            pXLColumn[43] = 11;  // 주(현) 임원퇴직소득금액 한도초과액
            pXLColumn[44] = 17;  // 종(전)1임원퇴직소득금액 한도초과액
            pXLColumn[45] = 23;  // 종(전)2임원퇴직소득금액 한도초과액

            pXLColumn[46] = 11;  // 주(현) 계
            pXLColumn[47] = 17;  // 종(전)1계
            pXLColumn[48] = 23;  // 종(전)2계

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 5. 비과세 및 감면소득 명세
            // -----------------------------------------------------------------------------------------------------------------------------------
            pXLColumn[49] = 11;  // 비과세_주(현)국외근로
            pXLColumn[50] = 17;  // 비과세_종(전)1국외근로
            pXLColumn[51] = 23;  // 비과세_종(전)2국외근로

            pXLColumn[52] = 11;  // 비과세_주(현)야간근로수당 
            pXLColumn[53] = 17;  // 비과세_종(전)1야간근로수당
            pXLColumn[54] = 23;  // 비과세_종(전)2야간근로수당

            pXLColumn[55] = 11;  // 비과세_주(현)출산/보육수당
            pXLColumn[56] = 17;  // 비과세_종(전)1출산/보육수당
            pXLColumn[57] = 23;  // 비과세_종(전)2출산/보육수당

            pXLColumn[58] = 11;  // 비과세_주(현)연구보조비
            pXLColumn[59] = 17;  // 비과세_종(전)1연구보조비
            pXLColumn[60] = 23;  // 비과세_종(전)2연구보조비

            pXLColumn[61] = 11;  // 비과세_주(현)수련보조수당
            pXLColumn[62] = 17;  // 비과세_종(전)1수련보조수당
            pXLColumn[63] = 23;  // 비과세_종(전)2수련보조수당

            pXLColumn[64] = 11;  // 비과세_주(현)비과세소득 계
            pXLColumn[65] = 17;  // 비과세_종(전)1비과세소득 계
            pXLColumn[66] = 23;  // 비과세_종(전)2비과세소득 계

            pXLColumn[67] = 11;  // 비과세_주(현)감면소득 계
            pXLColumn[68] = 17;  // 비과세_종(전)1감면소득 계
            pXLColumn[69] = 23;  // 비과세_종(전)2감면소득 계

            //--------------------------------------------------------------------------------------------------------------------
            // 6. 세액 명세
            //--------------------------------------------------------------------------------------------------------------------
            pXLColumn[70] = 19;  // 결정세액_소득세  
            pXLColumn[71] = 27;  // 결정세액_지방소득세    
            pXLColumn[72] = 36;  // 결정세액_농특세  

            pXLColumn[73] = 14;  // 기납부세액_종(전)1사업자번호  
            pXLColumn[74] = 19;  // 기납부세액_종(전)1소득세  
            pXLColumn[75] = 27;  // 기납부세액_종(전)1지방소득세      
            pXLColumn[76] = 36;  // 기납부세액_종(전)1농특세   

            pXLColumn[77] = 14;  // 기납부세액_종(전)2사업자번호    
            pXLColumn[78] = 19;  // 기납부세액_종(전)2소득세       
            pXLColumn[79] = 27;  // 기납부세액_종(전)2지방소득세      
            pXLColumn[80] = 36;  // 기납부세액_종(전)2농특세  

            pXLColumn[81] = 19;  // 기납부세액_주(현)소득세  
            pXLColumn[82] = 27;  // 기납부세액_주(현)지방소득세         
            pXLColumn[83] = 36;  // 기납부세액_주(현)농특세  

            pXLColumn[81] = 19;  // 기납부세액_주(현)소득세  
            pXLColumn[82] = 27;  // 기납부세액_주(현)지방소득세         
            pXLColumn[83] = 36;  // 기납부세액_주(현)농특세  

            pXLColumn[84] = 19;  // 차감징수세액_소득세 
            pXLColumn[85] = 27;  // 차감징수세액_지방소득세       
            pXLColumn[86] = 36;  // 차감징수세액_농특세 

            //--------------------------------------------------------------------------------------------------------------------
            //[ 2 page ]
            //--------------------------------------------------------------------------------------------------------------------
            pXLColumn[87] = 17;  // 총급여
            pXLColumn[88] = 17;  // 근로소득공제    
            pXLColumn[89] = 17;  // 근로소득금액

            // 기본공제
            pXLColumn[90] = 17;  // 기본(본인)
            pXLColumn[91] = 17;  // 기본(배우자)
            pXLColumn[92] = 10;  // 기본(부양인원 - 인원)       
            pXLColumn[93] = 17;  // 기본(부양인원 - 금액)

            // 추가공제
            pXLColumn[94] = 10;  // 추가공제(경로수 - 인원) 
            pXLColumn[95] = 17;  // 추가공제(경로수 - 금액)
            pXLColumn[96] = 10;  // 추가공제(장애인 - 인원)           
            pXLColumn[97] = 17;  // 추가공제(장애인 - 금액)
            pXLColumn[98] = 17;  // 추가공제(부녀세대)
            pXLColumn[99] = 10;  // 추가공제(자녀양육 - 인원)        
            pXLColumn[100] = 17;  // 추가공제(자녀양육 - 금액)
            pXLColumn[101] = 11;  // 추가공제(출산입양 - 인원)        
            pXLColumn[102] = 17;  // 추가공제(출산입양 - 금액)
            pXLColumn[103] = 17;  // 추가공제(한부모가족)

            pXLColumn[104] = 10;  // 다자녀공제(인원)  
            pXLColumn[105] = 17;  // 다자녀공제(금액) 

            // 연금보험료공제 
            pXLColumn[106] = 17;  // 국민연금보험료공제  

            pXLColumn[107] = 17;  // 공무원 연금 
            pXLColumn[108] = 17;  // 군인연금
            pXLColumn[109] = 17;  // 사립학교 교직원 연금
            pXLColumn[110] = 17;  // 별정우체국 연금

            pXLColumn[111] = 17;  // 과학기술인공제
            pXLColumn[112] = 17;  // 근로자퇴직급여 보장법에 따른 퇴직연금
            pXLColumn[113] = 17;  // 연금저축

            //특별소득공제
            pXLColumn[114] = 17;  // 건강보험료   
            pXLColumn[115] = 17;  // 고용보험료   
            pXLColumn[116] = 17;  // 보장성보험     
            pXLColumn[117] = 17;  // 장애인전용   

            pXLColumn[118] = 17;  // 의료비 (장애인)
            pXLColumn[119] = 17;  // 의료비 (기타)

            pXLColumn[120] = 17;  // 교육비 (장애인)
            pXLColumn[121] = 17;  // 교육비 (기타)

            pXLColumn[122] = 17;  // 주택임차차입금원리금상환액 (대출기관) 
            pXLColumn[123] = 17;  // 주택임차차입금원리금상환액 (거주자)

            pXLColumn[124] = 17;  // 월세액

            pXLColumn[125] = 17;  // 장기주택저당차입금이자상환액 - 2011이전 (15년미만)
            pXLColumn[126] = 17;  // 장기주택저당차입금이자상환액 - 2011이전 (15년~29년)
            pXLColumn[127] = 17;  // 장기주택저당차입금이자상환액 - 2011이전 (30년 이상)

            pXLColumn[128] = 17;  // 장기주택저당차입금이자상환액 - 2012이후 (고정금리비거치 상환대출)
            pXLColumn[129] = 17;  // 장기주택저당차입금이자상환액 - 2012이후 (기타대출)

            pXLColumn[130] = 17;  // 정치자금기부금 
            pXLColumn[131] = 17;  // 법정기부금 
            pXLColumn[132] = 17;  // 우리사주조합기부금  
            pXLColumn[133] = 17;  // 지정기부금

            pXLColumn[134] = 17;  // 계

            pXLColumn[135] = 17;  // 표준공제

            pXLColumn[136] = 17;  // 차감소득금액

            //그밖의소득공제
            pXLColumn[137] = 37;  // 개인연금저축소득공제 

            pXLColumn[138] = 37;  // 소기업/소상공인 공제부금 소득공제

            pXLColumn[139] = 37;  // 청약저축
            pXLColumn[140] = 37;  // 주택청약종합저축
            pXLColumn[141] = 37;  // 근로자주택마련저축

            pXLColumn[142] = 37;  // 투자조합출자등 소득공제
            pXLColumn[143] = 37;  // 신용카드등소득공제
            pXLColumn[144] = 37;  // 우리사주조합소득공제
            pXLColumn[145] = 37;  // 우리사주조합소득공제
            pXLColumn[146] = 37;  // 고용유지중소기업소득공제

            pXLColumn[147] = 37;  // 그 밖의 소득공제 계   

            pXLColumn[148] = 37;  // 특별공제 종합한도 초과액 

            pXLColumn[149] = 37;  // 종합과세표준

            pXLColumn[150] = 37;  // 산출세액

            //세액감면
            pXLColumn[151] = 37;  // 소득세법
            pXLColumn[152] = 37;  // 조세특례제한법 <53>-1 제외 
            pXLColumn[153] = 37;  // 조세특례제한법 제30조 
            pXLColumn[154] = 37;  // 조세조약

            pXLColumn[155] = 37;  // 세액감면 계
            pXLColumn[156] = 37;  // 근로소득
            pXLColumn[157] = 37;  // 납세조합공제
            pXLColumn[158] = 37;  // 주택차입금
            pXLColumn[159] = 37;  // 기부 정치자금
            pXLColumn[160] = 37;  // 외국 납부

            pXLColumn[161] = 37;  // 세액공제 계    

            //결정세액
            pXLColumn[162] = 37;  // 결정세액

            //그밖의것들
            pXLColumn[163] = 37;  // 대표이사

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 추가 종(전)3 
            // ----------------------------------------------------------------------------------------------------------------------------------- 
            pXLColumn[164] = 29;  // 종(전)
            pXLColumn[165] = 29;  // 종(전)3근무처명
            pXLColumn[166] = 29;  // 종(전)3사업자번호 
            pXLColumn[167] = 29;  // 종(전)3근무기간
            pXLColumn[168] = 29;  // 종(전)3감면기간
            pXLColumn[169] = 29;  // 종(전)3급여
            pXLColumn[170] = 29;  // 종(전)3상여
            pXLColumn[171] = 29;  // 종(전)3인정상여
            pXLColumn[172] = 29;  // 종(전)3주식매수선택권
            pXLColumn[173] = 29;  // 종(전)3우리사주조합인출금
            pXLColumn[174] = 29;  // 종(전)3임원퇴직소득금액 한도초과액
            pXLColumn[175] = 29;  // 종(전)3계 

            pXLColumn[176] = 29;  // 비과세_종(전)3국외근로
            pXLColumn[177] = 29;  // 비과세_종(전)3야간근로수당
            pXLColumn[178] = 29;  // 비과세_종(전)3출산/보육수당
            pXLColumn[179] = 29;  // 비과세_종(전)3연구보조비
            pXLColumn[180] = 29;  // 비과세_종(전)3수련보조수당
            pXLColumn[181] = 29;  // 비과세_종(전)3출산/보육수당
            pXLColumn[182] = 29;  // 비과세_종(전)3감면소득 계

            pXLColumn[183] = 14;  // 기납부세액_종(전)3사업자번호   
            pXLColumn[184] = 19;  // 기납부세액_종(전)3소득세   
            pXLColumn[185] = 27;  // 기납부세액_종(전)3지방소득세    
            pXLColumn[186] = 36;  // 기납부세액_종(전)3농특세   

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 추가 종(전)중소기업에 취업하는 청년에 대한 소득세 감면
            // ----------------------------------------------------------------------------------------------------------------------------------- 
            pXLColumn[187] = 2;   // 중소기업에 취업하는 청년에 대한 소득세 감면 존재
            pXLColumn[188] = 17;  // 중소기업에 취업하는 청년에 대한 소득세 감면1
            pXLColumn[189] = 23;  // 중소기업에 취업하는 청년에 대한 소득세 감면2 
            pXLColumn[190] = 29;  // 중소기업에 취업하는 청년에 대한 소득세 감면3
        }

        #endregion;

        #region ----- Array Set 6 ----

        private void SetArray6(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_SUPPORT_FAMILY, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[36];
            pXLColumn = new int[36];

            pGDColumn[0] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("MANY_CHILD_DED_COUNT");  // 다자녀 인원수
            pGDColumn[1] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("RELATION_CODE");         // 관계코드         
            pGDColumn[2] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("FAMILY_NAME");           // 성명       

            pGDColumn[3] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BASE_YN");               // 기본공제         
            pGDColumn[4] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("OLD_YN");                // 경로우대         
            pGDColumn[5] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BIRTH_YN");              // 출산/입양양육    

            pGDColumn[6] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("INSURE_AMT");            // 국세청-보험료    
            pGDColumn[7] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("MEDICAL_AMT");           // 국세청-의료비    
            pGDColumn[8] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("EDU_AMT");               // 국세청-교육비    
            pGDColumn[9] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CREDIT_AMT");            // 국세청-신용카드  
            pGDColumn[10] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHECK_CREDIT_AMT");      // 국세청-직불카드  
            pGDColumn[11] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CASH_AMT");              // 국세청-현금  
            pGDColumn[12] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("TRAD_MARKET_AMT");       // 국세청-전통시장   
            pGDColumn[13] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("PUBLIC_TRANSIT_AMT");       // 국세청-대중교통 
            pGDColumn[14] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DONAT_AMT");             // 국세청-기부금  

            pGDColumn[15] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("NATIONALITY_TYPE");      // 국가타입         
            pGDColumn[16] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("REPRE_NUM");             // 주민번호 

            pGDColumn[17] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("WOMAN_YN");              // 부녀자
            pGDColumn[18] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("SINGLE_PARENT_DED_YN");  // 한부모
            pGDColumn[19] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DISABILITY_YN");         // 장애인           
            pGDColumn[20] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHILD_YN");              // 자녀양육(6세이하)

            pGDColumn[21] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_INSURE_AMT");        // 기타-보험료      
            pGDColumn[22] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_MEDICAL_AMT");       // 기타-의료비      
            pGDColumn[23] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_EDU_AMT");           // 기타-교육비      
            pGDColumn[24] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_CREDIT_AMT");        // 기타-신용카드    
            pGDColumn[25] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHECK_ETC_CREDIT_AMT");  // 기타-직불카드    
            pGDColumn[26] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_CASH_AMT");          // 기타-현금        
            pGDColumn[27] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_TRAD_MARKET_AMT");   // 기타-전통시장  
            pGDColumn[28] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_PUBLIC_TRANSIT_AMT");    // 기타-대중교통  
            pGDColumn[29] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_DONAT_AMT");         // 기타-기부금 

            pGDColumn[30] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BASE_COUNT");            // 기본공제
            pGDColumn[31] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("OLD_COUNT");             // 경로우대
            pGDColumn[32] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BIRTH_COUNT");           // 출산/입양양육
            pGDColumn[33] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DISABILITY_COUNT");      // 장애인
            pGDColumn[34] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHILD_COUNT");           // 자녀양육
            pGDColumn[35] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("WOMAN_COUNT");           // 부녀세대

            //---------------------------------------------------------------------------------------------------------------
            pXLColumn[0] = 7;   // 다자녀 인원수
            pXLColumn[1] = 1;   // 관계코드         
            pXLColumn[2] = 3;   // 성명     

            pXLColumn[3] = 9;   // 기본공제         
            pXLColumn[4] = 11;  // 경로우대         
            pXLColumn[5] = 13;  // 출산/입양양육 

            pXLColumn[6] = 17;  // 국세청-보험료   
            pXLColumn[7] = 20;  // 국세청-의료비  
            pXLColumn[8] = 23;  // 국세청-교육비

            pXLColumn[9] = 26;  // 국세청-신용카드     
            pXLColumn[10] = 29; // 국세청-직불카드  
            pXLColumn[11] = 32; // 국세청-현금영수증  
            pXLColumn[12] = 35; // 국세청-전통시장이용액
            pXLColumn[13] = 38; // 국세청-대중교통이용액
            pXLColumn[14] = 41; // 국세청-기부금

            pXLColumn[15] = 1;  // 국세청-국가타입
            pXLColumn[16] = 3;  // 국세청-주민번호

            pXLColumn[17] = 9;  // 부녀자           
            pXLColumn[18] = 10;  // 한부모
            pXLColumn[19] = 11;  // 장애인
            pXLColumn[20] = 13;  // 6세이하

            pXLColumn[21] = 17;  // 기타-보험료      
            pXLColumn[22] = 20;  // 기타-의료비      
            pXLColumn[23] = 23;  // 기타-교육비    

            pXLColumn[24] = 26;  // 국세청-신용카드     
            pXLColumn[25] = 29; // 국세청-직불카드  
            pXLColumn[26] = 32; // 국세청-현금영수증  
            pXLColumn[27] = 35; // 국세청-전통시장이용액
            pXLColumn[28] = 38; // 국세청-대중교통이용액
            pXLColumn[29] = 41; // 국세청-기부금
        }

        #endregion;

        #region ----- Array Set 7 ----

        private void SetArray7(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[234];
            pXLColumn = new int[234];

            //-----------------------------------------------------------------------------------------------------------------------------------
            //-- 1. 오른쪽 상단 표 
            //-----------------------------------------------------------------------------------------------------------------------------------
            pGDColumn[0] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TYPE");        // 거주 구분(거주자1/거주자2)  

            pGDColumn[1] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_NAME");        // 거주지국
            pGDColumn[2] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_CODE");        // 거주지국코드

            pGDColumn[3] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONALITY_TYPE");     // 내외국인 구분(내국인1/외국인9)

            pGDColumn[4] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FOREIGN_TAX_YN");       // 외국인단일세율적용

            pGDColumn[5] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONAL_NAME");        // 국적
            pGDColumn[6] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONAL_CODE");        // 국적코드

            pGDColumn[7] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSEHOLD_TYPE");       // 세대주 구분(세대주1/세대원2)    

            pGDColumn[8] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_KEEP_TYPE");       // 연말정산구분


            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 2. 징수 의무자
            //-----------------------------------------------------------------------------------------------------------------------------------  
            pGDColumn[9] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CORP_NAME");            // 법인명(상호)                  
            pGDColumn[10] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRESIDENT_NAME");      // 대표자(성명)      

            pGDColumn[11] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("VAT_NUMBER");          // 사업자등록번호   

            pGDColumn[12] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ORG_ADDRESS");         // 소재지(주소)

            //-----------------------------------------------------------------------------------------------------------------------------------
            // --3.소득자
            // ---------------------------------------------------------------------------------------------------------------------------------- -  
            pGDColumn[13] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NAME");                // 성명
            pGDColumn[14] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPRE_NUM");           // 주민번호

            pGDColumn[15] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PERSON_ADDRESS");      // 주소  

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 4. 근무처별소득명세 : 
            // -----------------------------------------------------------------------------------------------------------------------------------        
            pGDColumn[16] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME");      // 주(현)근무처명
            pGDColumn[17] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NAME1");    // 종(전)1근무처명 
            pGDColumn[18] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NAME2");    // 종(전)2근무처명

            pGDColumn[19] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER");     // 주(현)사업자번호 
            pGDColumn[20] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NUM1");     // 종(전)1사업자번호
            pGDColumn[21] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NUM2");     // 종(전)2사업자번호 

            pGDColumn[22] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE");         // 주(현)근무기간
            pGDColumn[23] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE1");        // 종(전)1근무기간
            pGDColumn[24] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE2");        // 종(전)2근무기간

            pGDColumn[25] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE");         // 주(현)감면기간 
            pGDColumn[26] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE1");        // 종(전)1감면기간
            pGDColumn[27] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE2");        // 종(전)2감면기간

            pGDColumn[28] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_PAY_TOT_AMT");     // 주(현)급여 
            pGDColumn[29] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PAY_TOTAL_AMT1");      // 종(전)1급여
            pGDColumn[30] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PAY_TOTAL_AMT2");      // 종(전)2급여

            pGDColumn[31] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_BONUS_TOT_AMT");   // 주(현)상여   
            pGDColumn[32] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BONUS_TOTAL_AMT1");    // 종(전)1상여 
            pGDColumn[33] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BONUS_TOTAL_AMT2");    // 종(전)2상여

            pGDColumn[34] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_ADD_BONUS_AMT");   // 주(현)인정상여   
            pGDColumn[35] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_BONUS_AMT1");      // 종(전)1인정상여
            pGDColumn[36] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_BONUS_AMT2");      // 종(전)2인정상여

            pGDColumn[37] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_STOCK_BENE_AMT");  // 주(현)주식매수선택권
            pGDColumn[38] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STOCK_BENE_AMT1");     // 종(전)1주식매수선택권
            pGDColumn[39] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STOCK_BENE_AMT2");     // 종(전)2주식매수선택권

            pGDColumn[40] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_EMPLOYEE_STOCK_AMT");       // 주(현)우리사주조합인출금
            pGDColumn[41] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPLOYEE_STOCK_AMT1");       // 종(전)1우리사주조합인출금
            pGDColumn[42] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPLOYEE_STOCK_AMT2");               // 종(전)2우리사주조합인출금

            pGDColumn[43] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_OFFICE_RETIRE_OVER_AMT");    // 주(현)임원퇴직소득금액 한도초과액
            pGDColumn[44] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OFFICE_RETIRE_OVER_AMT1");   // 종(전)1임원퇴직소득금액 한도초과액
            pGDColumn[45] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OFFICE_RETIRE_OVER_AMT2");   // 종(전)2임원퇴직소득금액 한도초과액

            pGDColumn[46] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NOW_TOTAL_AMOUNT");    // 주(현)계     
            pGDColumn[47] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AMOUNT1");       // 종(전)1계   
            pGDColumn[48] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AMOUNT2");       // 종(전)2계  

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 5. 비과세 및 감면소득 명세
            // -----------------------------------------------------------------------------------------------------------------------------------
            pGDColumn[49] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_OUTSIDE_AMT");  // 비과세_주(현)국외근로
            pGDColumn[50] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OUTSIDE_AMT1");     // 비과세_종(전)1국외근로
            pGDColumn[51] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OUTSIDE_AMT2");     // 비과세_종(전)2국외근로

            pGDColumn[52] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_OT_AMT");       // 비과세_주(현)야간근로수당 
            pGDColumn[53] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OT_AMT1");          // 비과세_종(전)1야간근로수당
            pGDColumn[54] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OT_AMT2");          // 비과세_종(전)2야간근로수당

            pGDColumn[55] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_BIRTH_AMT");    // 비과세_주(현)출산/보육수당
            pGDColumn[56] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_BIRTH_AMT1");       // 비과세_종(전)1출산/보육수당
            pGDColumn[57] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_BIRTH_AMT2");       // 비과세_종(전)2출산/보육수당

            pGDColumn[58] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_COMPANY_AMT");   // 비과세_주(현)연구보조비
            pGDColumn[59] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_COMPANY_AMT1");  // 비과세_종(전)1연구보조비
            pGDColumn[60] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_COMPANY_AMT2");  // 비과세_종(전)2연구보조비

            pGDColumn[61] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TRAIN_AMT");     // 비과세_주(현)수련보조수당
            pGDColumn[62] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TRAIN_AMT1");    // 비과세_종(전)1수련보조수당
            pGDColumn[63] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TRAIN_AMT2");    // 비과세_종(전)2수련보조수당

            pGDColumn[64] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TOTAL_AMOUNT");  // 비과세_주(현)비과세소득 계
            pGDColumn[65] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_TOTAL_AMOUNT1");     // 비과세_종(전)1비과세소득 계
            pGDColumn[66] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_TOTAL_AMOUNT2");     // 비과세_종(전)2비과세소득 계

            pGDColumn[67] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT");  // 비과세_주(현)감면소득 계
            pGDColumn[68] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT1"); // 비과세_종(전)1감면소득 계
            pGDColumn[69] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT2"); // 비과세_종(전)2감면소득 계

            //--------------------------------------------------------------------------------------------------------------------
            // 6. 세액 명세
            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[70] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_IN_TAX_AMT");      // 결정세액_소득세               
            pGDColumn[71] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_LOCAL_TAX_AMT");   // 결정세액_지방소득세               
            pGDColumn[72] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_SP_TAX_AMT");      // 결정세액_농특세               

            pGDColumn[73] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_COMPANY_NUM1");    // 기납부세액_종(전)1사업자번호  
            pGDColumn[74] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_IN_TAX_AMT1");     // 기납부세액_종(전)1소득세      
            pGDColumn[75] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_LOCAL_TAX_AMT1");  // 기납부세액_종(전)1지방소득세      
            pGDColumn[76] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW1_SP_TAX_AMT1");     // 기납부세액_종(전)1농특세      


            pGDColumn[77] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_COMPANY_NUM2");    // 기납부세액_종(전)2사업자번호  
            pGDColumn[78] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_IN_TAX_AMT2");     // 기납부세액_종(전)2소득세      
            pGDColumn[79] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_LOCAL_TAX_AMT2");  // 기납부세액_종(전)2지방소득세      
            pGDColumn[80] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW2_SP_TAX_AMT2");     // 기납부세액_종(전)2농특세      

            pGDColumn[81] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_IN_TAX_AMT");      // 기납부세액_주(현)소득세       
            pGDColumn[82] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LOCAL_TAX_AMT");   // 기납부세액_주(현)지방소득세       
            pGDColumn[83] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_SP_TAX_AMT");      // 기납부세액_주(현)농특세       

            pGDColumn[84] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_IN_TAX_AMT");     // 차감징수세액_소득세           
            pGDColumn[85] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_LOCAL_TAX_AMT");  // 차감징수세액_지방소득세         
            pGDColumn[86] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_SP_TAX_AMT");     // 차감징수세액_농특세           


            //--------------------------------------------------------------------------------------------------------------------
            //[ 2 page ]
            //--------------------------------------------------------------------------------------------------------------------

            pGDColumn[87] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_TOT_AMT");           // 총급여
            pGDColumn[88] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_DED_AMT");           // 근로소득공제
            pGDColumn[89] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_AMT");               // 근로소득금액

            // 기본공제
            pGDColumn[90] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PER_DED_AMT");              // 기본(본인)
            pGDColumn[91] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SPOUSE_DED_AMT");           // 기본(배우자)
            pGDColumn[92] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUPP_DED_COUNT");           // 기본(부양인원 - 인원)            
            pGDColumn[93] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUPP_DED_AMT");             // 기본(부양인원 - 금액)

            // 추가공제
            pGDColumn[94] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OLD_DED_COUNT");            // 추가공제(경로수 - 인원)          
            pGDColumn[95] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OLD_DED_AMT");              // 추가공제(경로수 - 금액)
            pGDColumn[96] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_DED_COUNT");     // 추가공제(장애인 - 인원)          
            pGDColumn[97] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_DED_AMT");       // 추가공제(장애인 - 금액)
            pGDColumn[98] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WOMAN_DED_AMT");            // 추가공제(부녀세대)
            pGDColumn[99] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CHILD_DED_COUNT");          // 추가공제(자녀양육 - 인원)        
            pGDColumn[100] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CHILD_DED_AMT");            // 추가공제(자녀양육 - 금액)
            pGDColumn[101] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BIRTH_DED_COUNT");          // 추가공제(출산입양 - 인원)        
            pGDColumn[102] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BIRTH_DED_AMT");           // 추가공제(출산입양 - 금액)
            pGDColumn[103] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SINGLE_PARENT_DED_AMT");   // 추가공제(한부모가족)

            pGDColumn[104] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MANY_CHILD_DED_COUNT");    // 다자녀공제(인원)                 
            pGDColumn[105] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MANY_CHILD_DED_AMT");      // 다자녀공제(금액) 

            // 연금보험료공제 
            pGDColumn[106] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATI_ANNU_AMT");           // 국민연금보험료공제               

            pGDColumn[107] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PUBLIC_INSUR_AMT");        // 공무원 연금  
            pGDColumn[108] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MARINE_INSUR_AMT");        // 군인연금
            pGDColumn[109] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SCHOOL_STAFF_INSUR_AMT");  // 사립학교 교직원 연금
            pGDColumn[110] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("POST_OFFICE_INSUR_AMT");   // 별정우체국 연금

            pGDColumn[111] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SCIENTIST_ANNU_AMT");      // 과학기술인공제
            pGDColumn[112] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETR_ANNU_AMT");           // 근로자퇴직급여 보장법에 따른 퇴직연금
            pGDColumn[113] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ANNU_BANK_AMT");           // 연금저축

            //특별소득공제
            pGDColumn[114] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MEDIC_INSUR_AMT");         // 건강보험료                       
            pGDColumn[115] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HIRE_INSUR_AMT");          // 고용보험료                       
            pGDColumn[116] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("GUAR_INSUR_AMT");          // 보장성보험                       
            pGDColumn[117] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_INSUR_AMT");    // 장애인전용 

            pGDColumn[118] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_MEDIC_AMT");    // 의료비 (장애인)              
            pGDColumn[119] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ETC_MEDIC_AMT");           // 의료비 (기타)

            pGDColumn[120] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DISABILITY_EDUCATION_AMT");// 교육비 (장애인)                          
            pGDColumn[121] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ETC_EDUCATION_AMT");       // 교육비 (기타)

            pGDColumn[122] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_INTER_AMT");         // 주택임차차입금원리금상환액 (대출기관) 
            pGDColumn[123] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_INTER_AMT_ETC");     // 주택임차차입금원리금상환액 (거주자)

            pGDColumn[124] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_MONTHLY_AMT");       // 월세액

            pGDColumn[125] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT");     // 장기주택저당차입금이자상환액 - 2011이전 (15년미만)
            pGDColumn[126] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_1");   // 장기주택저당차입금이자상환액 - 2011이전 (15년~29년)
            pGDColumn[127] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_2");   // 장기주택저당차입금이자상환액 - 2011이전 (30년 이상)

            pGDColumn[128] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_3_FIX"); // 장기주택저당차입금이자상환액 - 2012이후 (고정금리비거치 상환대출)
            pGDColumn[129] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_HOUSE_PROF_AMT_3_ETC"); // 장기주택저당차입금이자상환액 - 2012이후 (기타)

            pGDColumn[130] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_DONAT_POLI_AMT");  // 정치자금기부금                       
            pGDColumn[131] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DONAT_DED_ALL");           // 법정기부금                       
            pGDColumn[132] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DONAT_DED_30");            // 우리사주조합기부금                       
            pGDColumn[133] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DONAT_DED");               // 지정기부금 

            pGDColumn[134] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SP_DED_SUM");              // 계

            pGDColumn[135] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STAND_DED_AMT");           // 표준공제

            //차감소득금액
            pGDColumn[136] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SUBT_DED_AMT");            // 차감소득금액


            //그밖의소득공제

            pGDColumn[137] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PERS_ANNU_BANK_AMT");      //개인연금저축소득공제 

            pGDColumn[138] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SMALL_CORPOR_DED_AMT");    // 소기업/소상공인 공제부금 소득공제

            pGDColumn[139] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_APP_SAVE_AMT");      // 청약저축
            pGDColumn[140] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HOUSE_APP_DEPOSIT_AMT");   // 주택청약종합저축
            pGDColumn[141] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORKER_HOUSE_SAVE_AMT");   // 근로자주택마련저축

            pGDColumn[142] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INVES_AMT");               // 투자조합출자등 소득공제
            pGDColumn[143] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CREDIT_AMT");              // 신용카드등소득공제
            pGDColumn[144] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPL_STOCK_AMT");          // 우리사주조합소득공제
            pGDColumn[145] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HIRE_KEEP_EMPLOY_AMT");    // 고용유지중소기업근로자
            pGDColumn[146] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FIX_LEASE_DED_AMT");       // 목돈안드는전세이자상환

            pGDColumn[147] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ETC_DED_SUM");             // 그 밖의 소득공제 계   

            pGDColumn[148] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SP_DED_TOT_AMT");          // 특별공제 종합한도 초과액 

            pGDColumn[149] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_STD_AMT");             // 종합과세표준

            pGDColumn[150] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("COMP_TAX_AMT");            // 산출세액

            //세액감면

            pGDColumn[151] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_IN_LAW_AMT");     // 소득세법
            pGDColumn[152] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SP_LAW_AMT");     // 조세특례제한법 <53>-1 제외 
            pGDColumn[153] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SP_LAW_AMT2");    // 조세특례제한법 제30조 
            pGDColumn[154] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_LAW_AMT");        // 조세조약

            pGDColumn[155] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SUM");            // 세액감면 계
            pGDColumn[156] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_INCOME_AMT");      // 근로소득
            pGDColumn[157] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_TAXGROUP_AMT");    // 납세조합공제
            pGDColumn[158] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_HOUSE_DEBT_AMT");  // 주택차입금
            pGDColumn[159] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_DONAT_POLI_AMT2"); // 기부 정치자금
            pGDColumn[160] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SUM");            // 외국 납부

            pGDColumn[161] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_SUM");             // 세액공제 계     

            //결정세액
            pGDColumn[162] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESULT_TAX_SUM");          // 결정세액

            //그밖의것들
            pGDColumn[163] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WITHHOLDING_OWNER");          // 대표이사


            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 추가 종(전)3 
            // ----------------------------------------------------------------------------------------------------------------------------------- 

            pGDColumn[164] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_EXIST3");         // 종(전)
            pGDColumn[165] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NAME3");          // 종(전)3근무처명
            pGDColumn[166] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW_COMPANY_NUM3");           // 종(전)3사업자번호 
            pGDColumn[167] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADJUST_DATE3");              // 종(전)3근무기간
            pGDColumn[168] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_DATE3");              // 종(전)3감면기간
            pGDColumn[169] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PAY_TOTAL_AMT3");            // 종(전)3급여
            pGDColumn[170] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("BONUS_TOTAL_AMT3");          // 종(전)3상여
            pGDColumn[171] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_BONUS_AMT3");            // 종(전)3인정상여
            pGDColumn[172] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("STOCK_BENE_AMT3");           // 종(전)3주식매수선택권
            pGDColumn[173] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EMPLOYEE_STOCK_AMT3");       // 종(전)3우리사주조합인출금
            pGDColumn[174] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OFFICE_RETIRE_OVER_AMT3");   // 종(전)3임원퇴직소득금액 한도초과액
            pGDColumn[175] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AMOUNT3");             // 종(전)3계 


            pGDColumn[176] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OUTSIDE_AMT3");          // 비과세_종(전)3국외근로
            pGDColumn[177] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_OT_AMT3");               // 비과세_종(전)3야간근로수당
            pGDColumn[178] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_BIRTH_AMT3");            // 비과세_종(전)3출산/보육수당
            pGDColumn[179] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_COMPANY_AMT3");      // 비과세_종(전)3연구보조비
            pGDColumn[180] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NONTAX_TRAIN_AMT3");        // 비과세_종(전)3수련보조수당
            pGDColumn[181] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NT_TOTAL_AMOUNT3");         // 비과세_종(전)3출산/보육수당
            pGDColumn[182] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REDUCE_TOTAL_AMOUNT3");     // 비과세_종(전)3감면소득 계

            pGDColumn[183] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW3_COMPANY_NUM3");        // 기납부세액_종(전)3사업자번호  
            pGDColumn[184] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW3_IN_TAX_AMT3");         // 기납부세액_종(전)3소득세      
            pGDColumn[185] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW3_LOCAL_TAX_AMT3");      // 기납부세액_종(전)3지방소득세      
            pGDColumn[186] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PW3_SP_TAX_AMT3");         // 기납부세액_종(전)3농특세   

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 추가 종(전)중소기업에 취업하는 청년에 대한 소득세 감면
            // -----------------------------------------------------------------------------------------------------------------------------------

            pGDColumn[187] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RD_SMALL_BUSINESS_AMT_EXIST");  // 중소기업에 취업하는 청년에 대한 소득세 감면 존재

            pGDColumn[188] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RD_SMALL_BUSINESS_AMT1");       // 중소기업에 취업하는 청년에 대한 소득세 감면1
            pGDColumn[189] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RD_SMALL_BUSINESS_AMT2");       // 중소기업에 취업하는 청년에 대한 소득세 감면2
            pGDColumn[190] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RD_SMALL_BUSINESS_AMT3");       // 중소기업에 취업하는 청년에 대한 소득세 감면3


            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 2014
            // -----------------------------------------------------------------------------------------------------------------------------------
            pGDColumn[191] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("FORWARD_DONATION_AMT");         // 2014-특별소득공제(기부금이월분)
            pGDColumn[192] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("COM_STOCK_DONATION_AMT");       // 2014-그밖의소득공제(우리사주조합기부금)
            pGDColumn[193] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INVEST_AMT_14");                // 2014-투자조합출자금액  14년
            pGDColumn[194] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_SET_INVEST_SAVING_AMT");   // 2014- 장기집합투자증권저축
            pGDColumn[195] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_REDU_SMALL_SP_LAW_AMT");    // 2014-세감 ( 중소기업취업청년)
            
            pGDColumn[196] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_CHILD_RAISE_DED_CNT");       // 2014-세공(자녀양육 인원)            
            pGDColumn[197] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_CHILD_RAISE_DED");          // 2014-세공(자녀양육 금액)
            pGDColumn[198] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_CHILD_6_UNDER_DED_CNT");     // 2014-세공(6세이하 자녀양육 인원)            
            pGDColumn[199] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_CHILD_6_UNDER_DED_AMT");     // 2014-세공(6세이하 자녀양육 금액)
            pGDColumn[200] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_BIRTH_DED_CNT");             // 2014-세공(출생입양 인원)            
            pGDColumn[201] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_BIRTH_DED_AMT");             // 2014-세공(출생입양 금액)
            pGDColumn[202] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_RETR_ANNU_AMT");        // 2014-연금계좌 합계 
            pGDColumn[203] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SEMA_SAVING_DED_AMOUNT");       // 2014-과학기술인 공제금액
            pGDColumn[204] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SEMA_SAVING_REAL_DED_AMT");     // 2014-과학기술인 세액공제
            pGDColumn[205] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETR_SAVING_DED_AMOUNT");       // 2014-퇴직연금  공제금액
            pGDColumn[206] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETR_SAVING_REAL_DED_AMT");     // 2014-퇴직연금 세액공제
            pGDColumn[207] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ANNU_SAVING_DED_AMOUNT");       // 2014-연금저축 공제금액
            pGDColumn[208] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ANNU_SAVING_REAL_DED_AMT");     // 2014-연금저축 세액공제
            pGDColumn[209] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_INSUR_TAR_AMT");        // 2014-세감 ( 보장성보험 공제금액)
            pGDColumn[210] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_INSUR_AMT");            // 2014-세감 ( 보장성보험 세액공제)
            pGDColumn[211] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_DISABILITY_INSUR_AMT");      // 2014-세감 ( 장애인보험 공제금액)
            pGDColumn[212] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_DISABILITY_INSUR_DED_AMT");  // 2014-세감 ( 장애인보험 세액공제)

            pGDColumn[213] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_MEDIC_TAR_AMT");        // 2014-세감 (의료비 공제금액)
            pGDColumn[214] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_MEDIC_AMT");            // 2014-세감 (의료비 세액공제)
            pGDColumn[215] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_EDUCATION_TAR_AMT");    // 2014-세감 (교육비 공제금액)
            pGDColumn[216] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_EDUCATION_AMT");        // 2014-세감 (교육비 세액공제) 
            pGDColumn[217] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_POLI_DONAT1");          // 2014-세감 (정치자금기부금-10만원이하) 공제금액

            pGDColumn[218] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_POLI_AMT1");            // 2014-세감 (정치자금기부금-10만원이하) 세액공제
            pGDColumn[219] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_POLI_DONAT2");          // 2014-세감 (정치자금기부금-10만원초과) 공제금액
            pGDColumn[220] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_POLI_AMT2");            // 2014-세감 (정치자금기부금-10만원이하) 세액공제
            pGDColumn[221] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_LEGAL_DONAT");          // 2014-세감 (법정기부금) 공제금액
            pGDColumn[222] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_LEGAL_AMT");            // 2014-세감 (정치자금기부금-10만원이하) 세액공제
            pGDColumn[223] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_DESIGN_DONAT");         // 2014-세감 (지정기부금) 공제금액
            pGDColumn[224] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_DESIGN_AMT");           // 2014-세감 (정치자금기부금-10만원이하) 세액공제
            pGDColumn[225] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ETC_DED_SUM_2014");             // 2014 그 밖의 소득공제 계 
            pGDColumn[226] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("SP_DED_SUM_2014");              // 2014-특별소득공제 합계
            pGDColumn[227] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_SPC_SUM_2014");         // 2014-특별세액공제 계   
            pGDColumn[228] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_SUM_2014");             // 2014-세액공제 계    

            pGDColumn[229] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_HOUSE_MONTHLY_AMT");          // 2014 세공 월세액 공제 대상금액
            pGDColumn[230] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TD_HOUSE_MONTHLY_DED_AMT");      //2014 세공 월세액 세액공제액

            pGDColumn[231] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECIPIENT_PERSON_NAME");           // 2014 소득자 보관
            pGDColumn[232] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECIPIENT_TAX_OFFICE_NAME");       // 2014 세무서 제출
            pGDColumn[233] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECIPIENT_ETC_NAME");              // 2014 발행자 보관


            //-----------------------------------------------------------------------------------------------------------------------------------
            //-- 1. 오른쪽 상단 표 
            //-----------------------------------------------------------------------------------------------------------------------------------
            pXLColumn[0] = 37;  // 거주 구분(거주자1/거주자2)  

            pXLColumn[1] = 31;  // 거주지국
            pXLColumn[2] = 39;  // 거주지국코드    

            pXLColumn[3] = 37;  // 내외국인 구분(내국인1/외국인9)

            pXLColumn[4] = 39;  // 외국인단일세율적용          

            pXLColumn[5] = 31;  // 국적                
            pXLColumn[6] = 39;  // 국적코드      

            pXLColumn[7] = 37;  // 세대주 구분(세대주1/세대원2)    

            pXLColumn[8] = 37;  // 연말정산구분           

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 2. 징수 의무자
            //-----------------------------------------------------------------------------------------------------------------------------------  
            pXLColumn[9] = 10;   // 법인명(상호)           
            pXLColumn[10] = 30;  // 대표자(성명)  

            pXLColumn[11] = 10;  // 사업자등록번호     

            pXLColumn[12] = 10;  // 소재지(주소)  

            //-----------------------------------------------------------------------------------------------------------------------------------
            // --3.소득자
            // ---------------------------------------------------------------------------------------------------------------------------------- -  
            pXLColumn[13] = 10;  // 성명
            pXLColumn[14] = 30;  // 주민번호

            pXLColumn[15] = 10;  // 주소 

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 4. 근무처별소득명세 : 
            // -----------------------------------------------------------------------------------------------------------------------------------  
            pXLColumn[16] = 10;  // 주(현)근무처명 
            pXLColumn[17] = 17;  // 종(전)1근무처명 
            pXLColumn[18] = 23;  // 종(전)2근무처명

            pXLColumn[19] = 10;  // 주(현)사업자번호 
            pXLColumn[20] = 17;  // 종(전)1사업자번호 
            pXLColumn[21] = 23;  // 종(전)2사업자번호 

            pXLColumn[22] = 10;  // 주(현)근무기간 
            pXLColumn[23] = 17;  // 종(전)1근무기간 
            pXLColumn[24] = 23;  // 종(전)2근무기간 

            pXLColumn[25] = 10;  // 주(현)감면기간 
            pXLColumn[26] = 17;  // 종(전)1감면기간 
            pXLColumn[27] = 23;  // 종(전)2감면기간 

            pXLColumn[28] = 10;  // 주(현)급여 
            pXLColumn[29] = 17;  // 종(전)1급여 
            pXLColumn[30] = 23;  // 종(전)2급여 

            pXLColumn[31] = 10;  // 주(현)상여 
            pXLColumn[32] = 17;  // 종(전)1상여 
            pXLColumn[33] = 23;  // 종(전)2상여 

            pXLColumn[34] = 10;  // 주(현) 인정상여 
            pXLColumn[35] = 17;  // 종(전)1인정상여 
            pXLColumn[36] = 23;  // 종(전)2인정상여 

            pXLColumn[37] = 11;  // 주(현) 주식매수선택권 
            pXLColumn[38] = 17;  // 종(전)1주식매수선택권 
            pXLColumn[39] = 23;  // 종(전)2주식매수선택권 

            pXLColumn[40] = 10;  // 주(현) 우리사주조합인출금 
            pXLColumn[41] = 17;  // 종(전)1우리사주조합인출금 
            pXLColumn[42] = 23;  // 종(전)2우리사주조합인출금 

            pXLColumn[43] = 10;  // 주(현) 임원퇴직소득금액 한도초과액
            pXLColumn[44] = 17;  // 종(전)1임원퇴직소득금액 한도초과액
            pXLColumn[45] = 23;  // 종(전)2임원퇴직소득금액 한도초과액

            pXLColumn[46] = 10;  // 주(현) 계
            pXLColumn[47] = 17;  // 종(전)1계
            pXLColumn[48] = 23;  // 종(전)2계

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 5. 비과세 및 감면소득 명세
            // -----------------------------------------------------------------------------------------------------------------------------------
            pXLColumn[49] = 10;  // 비과세_주(현)국외근로
            pXLColumn[50] = 17;  // 비과세_종(전)1국외근로
            pXLColumn[51] = 23;  // 비과세_종(전)2국외근로

            pXLColumn[52] = 10;  // 비과세_주(현)야간근로수당 
            pXLColumn[53] = 17;  // 비과세_종(전)1야간근로수당
            pXLColumn[54] = 23;  // 비과세_종(전)2야간근로수당

            pXLColumn[55] = 10;  // 비과세_주(현)출산/보육수당
            pXLColumn[56] = 17;  // 비과세_종(전)1출산/보육수당
            pXLColumn[57] = 23;  // 비과세_종(전)2출산/보육수당

            pXLColumn[58] = 10;  // 비과세_주(현)연구보조비
            pXLColumn[59] = 17;  // 비과세_종(전)1연구보조비
            pXLColumn[60] = 23;  // 비과세_종(전)2연구보조비

            pXLColumn[61] = 10;  // 비과세_주(현)수련보조수당
            pXLColumn[62] = 17;  // 비과세_종(전)1수련보조수당
            pXLColumn[63] = 23;  // 비과세_종(전)2수련보조수당

            pXLColumn[64] = 10;  // 비과세_주(현)비과세소득 계
            pXLColumn[65] = 17;  // 비과세_종(전)1비과세소득 계
            pXLColumn[66] = 23;  // 비과세_종(전)2비과세소득 계

            pXLColumn[67] = 10;  // 비과세_주(현)감면소득 계
            pXLColumn[68] = 17;  // 비과세_종(전)1감면소득 계
            pXLColumn[69] = 23;  // 비과세_종(전)2감면소득 계

            //--------------------------------------------------------------------------------------------------------------------
            // 6. 세액 명세
            //--------------------------------------------------------------------------------------------------------------------
            pXLColumn[70] = 19;  // 결정세액_소득세  
            pXLColumn[71] = 27;  // 결정세액_지방소득세    
            pXLColumn[72] = 36;  // 결정세액_농특세  

            pXLColumn[73] = 13;  // 기납부세액_종(전)1사업자번호  
            pXLColumn[74] = 19;  // 기납부세액_종(전)1소득세  
            pXLColumn[75] = 27;  // 기납부세액_종(전)1지방소득세      
            pXLColumn[76] = 36;  // 기납부세액_종(전)1농특세   

            pXLColumn[77] = 13;  // 기납부세액_종(전)2사업자번호    
            pXLColumn[78] = 19;  // 기납부세액_종(전)2소득세       
            pXLColumn[79] = 27;  // 기납부세액_종(전)2지방소득세      
            pXLColumn[80] = 36;  // 기납부세액_종(전)2농특세  

            pXLColumn[81] = 19;  // 기납부세액_주(현)소득세  
            pXLColumn[82] = 27;  // 기납부세액_주(현)지방소득세         
            pXLColumn[83] = 36;  // 기납부세액_주(현)농특세  

            pXLColumn[81] = 19;  // 기납부세액_주(현)소득세  
            pXLColumn[82] = 27;  // 기납부세액_주(현)지방소득세         
            pXLColumn[83] = 36;  // 기납부세액_주(현)농특세  

            pXLColumn[84] = 19;  // 차감징수세액_소득세 
            pXLColumn[85] = 27;  // 차감징수세액_지방소득세       
            pXLColumn[86] = 36;  // 차감징수세액_농특세 

            //--------------------------------------------------------------------------------------------------------------------
            //[ 2 page ]
            //--------------------------------------------------------------------------------------------------------------------
            pXLColumn[87] = 17;  // 총급여
            pXLColumn[88] = 17;  // 근로소득공제    
            pXLColumn[89] = 17;  // 근로소득금액

            // 기본공제
            pXLColumn[90] = 17;  // 기본(본인)
            pXLColumn[91] = 17;  // 기본(배우자)
            pXLColumn[92] = 9;  // 기본(부양인원 - 인원)       
            pXLColumn[93] = 17;  // 기본(부양인원 - 금액)

            // 추가공제
            pXLColumn[94] = 9;  // 추가공제(경로수 - 인원) 
            pXLColumn[95] = 17;  // 추가공제(경로수 - 금액)
            pXLColumn[96] = 9;  // 추가공제(장애인 - 인원)           
            pXLColumn[97] = 17;  // 추가공제(장애인 - 금액)
            pXLColumn[98] = 17;  // 추가공제(부녀세대)
            pXLColumn[99] = 9;  // 추가공제(자녀양육 - 인원)        
            pXLColumn[100] = 17;  // 추가공제(자녀양육 - 금액)
            pXLColumn[101] = 11;  // 추가공제(출산입양 - 인원)        
            pXLColumn[102] = 17;  // 추가공제(출산입양 - 금액)
            pXLColumn[103] = 17;  // 추가공제(한부모가족)

            pXLColumn[104] = 9;  // 다자녀공제(인원)  
            pXLColumn[105] = 17;  // 다자녀공제(금액) 

            // 연금보험료공제 
            pXLColumn[106] = 17;  // 국민연금보험료공제  

            pXLColumn[107] = 17;  // 공무원 연금 
            pXLColumn[108] = 17;  // 군인연금
            pXLColumn[109] = 17;  // 사립학교 교직원 연금
            pXLColumn[110] = 17;  // 별정우체국 연금

            pXLColumn[111] = 17;  // 과학기술인공제
            pXLColumn[112] = 17;  // 근로자퇴직급여 보장법에 따른 퇴직연금
            pXLColumn[113] = 17;  // 연금저축

            //특별소득공제
            pXLColumn[114] = 17;  // 건강보험료   
            pXLColumn[115] = 17;  // 고용보험료   
            pXLColumn[116] = 17;  // 보장성보험     
            pXLColumn[117] = 17;  // 장애인전용   

            pXLColumn[118] = 17;  // 의료비 (장애인)
            pXLColumn[119] = 17;  // 의료비 (기타)

            pXLColumn[120] = 17;  // 교육비 (장애인)
            pXLColumn[121] = 17;  // 교육비 (기타)

            pXLColumn[122] = 17;  // 주택임차차입금원리금상환액 (대출기관) 
            pXLColumn[123] = 17;  // 주택임차차입금원리금상환액 (거주자)

            pXLColumn[124] = 17;  // 월세액

            pXLColumn[125] = 17;  // 장기주택저당차입금이자상환액 - 2011이전 (15년미만)
            pXLColumn[126] = 17;  // 장기주택저당차입금이자상환액 - 2011이전 (15년~29년)
            pXLColumn[127] = 17;  // 장기주택저당차입금이자상환액 - 2011이전 (30년 이상)

            pXLColumn[128] = 17;  // 장기주택저당차입금이자상환액 - 2012이후 (고정금리비거치 상환대출)
            pXLColumn[129] = 17;  // 장기주택저당차입금이자상환액 - 2012이후 (기타대출)

            pXLColumn[130] = 17;  // 정치자금기부금 
            pXLColumn[131] = 17;  // 법정기부금 
            pXLColumn[132] = 17;  // 우리사주조합기부금  
            pXLColumn[133] = 17;  // 지정기부금

            pXLColumn[134] = 17;  // 계

            pXLColumn[135] = 37;  // 표준공제

            pXLColumn[136] = 17;  // 차감소득금액

            //그밖의소득공제
            pXLColumn[137] = 17;  // 개인연금저축소득공제 

            pXLColumn[138] = 17;  // 소기업/소상공인 공제부금 소득공제

            pXLColumn[139] = 17;  // 청약저축
            pXLColumn[140] = 17;  // 주택청약종합저축
            pXLColumn[141] = 17;  // 근로자주택마련저축

            pXLColumn[142] = 17;  // 투자조합출자등 소득공제
            pXLColumn[143] = 17;  // 신용카드등소득공제
            pXLColumn[144] = 17;  // 우리사주조합소득공제
            pXLColumn[145] = 17;  // 우리사주조합소득공제
            pXLColumn[146] = 17;  // 고용유지중소기업소득공제

            pXLColumn[147] = 17;  // 그 밖의 소득공제 계   

            pXLColumn[148] = 17;  // 특별공제 종합한도 초과액 

            pXLColumn[149] = 37;  // 종합과세표준

            pXLColumn[150] = 37;  // 산출세액

            //세액감면
            pXLColumn[151] = 37;  // 소득세법
            pXLColumn[152] = 37;  // 조세특례제한법 <53>-1 제외 
            pXLColumn[153] = 37;  // 조세특례제한법 제30조 
            pXLColumn[154] = 37;  // 조세조약

            pXLColumn[155] = 37;  // 세액감면 계
            pXLColumn[156] = 37;  // 근로소득
            pXLColumn[157] = 37;  // 납세조합공제
            pXLColumn[158] = 37;  // 주택차입금
            pXLColumn[159] = 37;  // 기부 정치자금
            pXLColumn[160] = 37;  // 외국 납부

            pXLColumn[161] = 37;  // 세액공제 계    

            //결정세액
            pXLColumn[162] = 37;  // 결정세액

            //그밖의것들
            pXLColumn[163] = 37;  // 대표이사

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 추가 종(전)3 
            // ----------------------------------------------------------------------------------------------------------------------------------- 
            pXLColumn[164] = 29;  // 종(전)
            pXLColumn[165] = 29;  // 종(전)3근무처명
            pXLColumn[166] = 29;  // 종(전)3사업자번호 
            pXLColumn[167] = 29;  // 종(전)3근무기간
            pXLColumn[168] = 29;  // 종(전)3감면기간
            pXLColumn[169] = 29;  // 종(전)3급여
            pXLColumn[170] = 29;  // 종(전)3상여
            pXLColumn[171] = 29;  // 종(전)3인정상여
            pXLColumn[172] = 29;  // 종(전)3주식매수선택권
            pXLColumn[173] = 29;  // 종(전)3우리사주조합인출금
            pXLColumn[174] = 29;  // 종(전)3임원퇴직소득금액 한도초과액
            pXLColumn[175] = 29;  // 종(전)3계 

            pXLColumn[176] = 29;  // 비과세_종(전)3국외근로
            pXLColumn[177] = 29;  // 비과세_종(전)3야간근로수당
            pXLColumn[178] = 29;  // 비과세_종(전)3출산/보육수당
            pXLColumn[179] = 29;  // 비과세_종(전)3연구보조비
            pXLColumn[180] = 29;  // 비과세_종(전)3수련보조수당
            pXLColumn[181] = 29;  // 비과세_종(전)3출산/보육수당
            pXLColumn[182] = 29;  // 비과세_종(전)3감면소득 계

            pXLColumn[183] = 14;  // 기납부세액_종(전)3사업자번호   
            pXLColumn[184] = 19;  // 기납부세액_종(전)3소득세   
            pXLColumn[185] = 27;  // 기납부세액_종(전)3지방소득세    
            pXLColumn[186] = 36;  // 기납부세액_종(전)3농특세   

            //-----------------------------------------------------------------------------------------------------------------------------------
            // -- 추가 종(전)중소기업에 취업하는 청년에 대한 소득세 감면
            // ----------------------------------------------------------------------------------------------------------------------------------- 
            pXLColumn[187] = 2;   // 중소기업에 취업하는 청년에 대한 소득세 감면 존재
            pXLColumn[188] = 17;  // 중소기업에 취업하는 청년에 대한 소득세 감면1
            pXLColumn[189] = 23;  // 중소기업에 취업하는 청년에 대한 소득세 감면2 
            pXLColumn[190] = 29;  // 중소기업에 취업하는 청년에 대한 소득세 감면3

            // 2014
            pXLColumn[191] = 17;    // 2014-특별소득공제(기부금이월분)
            pXLColumn[192] = 17;    // 2014-그밖의소득공제(우리사주조합기부금)
            pXLColumn[193] = 17;    // 2014-투자조합출자금액  14년
            pXLColumn[194] = 17;    // 2014- 장기집합투자증권저축
            pXLColumn[195] = 37;    // 2014-세감 ( 중소기업취업청년)

            pXLColumn[196] = 35;    // 2014-세공(자녀양육 인원)
            pXLColumn[197] = 37;    // 2014-세공(자녀양육 금액)

            pXLColumn[198] = 35;    // 2014-세공(6세이하 자녀양육 인원)
            pXLColumn[199] = 37;    // 2014-세공(6세이하 자녀양육 금액)
            pXLColumn[200] = 35;    // 2014-세공(출생입양 자녀양육 인원)
            pXLColumn[201] = 37;    // 2014-세공(출생입양 자녀양육 금액)

            pXLColumn[202] = 37;    // 2014-세감 ( 연금계좌) 
            pXLColumn[203] = 37;    // 2014-과학기술인 공제금액
            pXLColumn[204] = 37;    // 2014-과학기술인 세액공제
            pXLColumn[205] = 37;    // 2014-퇴직연금  공제금액
            pXLColumn[206] = 37;    // 2014-퇴직연금 세액공제
            pXLColumn[207] = 37;    // 2014-연금저축 공제금액
            pXLColumn[208] = 37;    // 2014-연금저축 세액공제
            pXLColumn[209] = 37;    // 2014-세감 ( 보장성보험 공제금액)
            pXLColumn[210] = 37;    // 2014-세감 ( 보장성보험 세액공제)
            pXLColumn[211] = 37;    // 2014-세감 ( 장애인보험 공제금액)
            pXLColumn[212] = 37;    // 2014-세감 ( 장애인보험 세액공제)

            pXLColumn[213] = 37;    // 2014-세감 (의료비 공제금액)
            pXLColumn[214] = 37;    // 2014-세감 (의료비 세액공제)
            pXLColumn[215] = 37;    // 2014-세감 (교육비 공제금액)
            pXLColumn[216] = 37;    // 2014-세감 (교육비 세액공제) 
            pXLColumn[217] = 37;    // 2014-세감 (정치자금기부금-10만원이하) 공제금액

            pXLColumn[218] = 37;    // 2014-세감 (정치자금기부금-10만원이하) 세액공제
            pXLColumn[219] = 37;    // 2014-세감 (정치자금기부금-10만원초과) 공제금액
            pXLColumn[220] = 37;    // 2014-세감 (정치자금기부금-10만원이하) 세액공제
            pXLColumn[221] = 37;    // 2014-세감 (법정기부금) 공제금액
            pXLColumn[222] = 37;    // 2014-세감 (정치자금기부금-10만원이하) 세액공제
            pXLColumn[223] = 37;    // 2014-세감 (지정기부금) 공제금액
            pXLColumn[224] = 37;    // 2014-세감 (정치자금기부금-10만원이하) 세액공제
            pXLColumn[225] = 17;    // 2014 그 밖의 소득공제 계 
            pXLColumn[226] = 17;    // 2014-특별소득공제 합계
            pXLColumn[227] = 37;    // 2014-특별세액공제 계   
            pXLColumn[228] = 37;    // 2014-세액공제 계  

            pXLColumn[229] = 37;    // 2014 세공 월세액 공제 대상금액
            pXLColumn[230] = 37;    // 2014 세공 월세액 세액공제액

            pXLColumn[231] = 2;     // 2014 소득자 보관
            pXLColumn[232] = 2;     // 2014 세무서 제출
            pXLColumn[233] = 2;     // 2014 발행자 보관
        }

        #endregion;

        #region ----- Array Set 8 ----

        private void SetArray8(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_SUPPORT_FAMILY, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[34];
            pXLColumn = new int[34];

            pGDColumn[0] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHILD_RAISE_COUNT");     // 다자녀 인원수
            pGDColumn[1] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("RELATION_CODE");         // 관계코드         
            pGDColumn[2] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("FAMILY_NAME");           // 성명       

            pGDColumn[3] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BASE_YN");               // 기본공제         
            pGDColumn[4] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("OLD_YN");                // 경로우대         
            pGDColumn[5] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BIRTH_YN");              // 출산/입양양육    

            pGDColumn[6] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("MEDIC_HIRE_INSUR_AMT");  // 국세청-건강/고용보험료 
            pGDColumn[7] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("INSURE_AMT");            // 국세청-보험료    
            pGDColumn[8] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DISABILITY_INSURE_AMT"); // 국세청-장애인보험료    
            pGDColumn[9] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("MEDICAL_AMT");           // 국세청-의료비    
            pGDColumn[10] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("EDU_AMT");               // 국세청-교육비    
            pGDColumn[11] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CREDIT_AMT");            // 국세청-신용카드  
            pGDColumn[12] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHECK_CREDIT_AMT");      // 국세청-직불카드  
            pGDColumn[13] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CASH_AMT");              // 국세청-현금  
            pGDColumn[14] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("TRAD_MARKET_AMT");       // 국세청-전통시장   
            pGDColumn[15] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("PUBLIC_TRANSIT_AMT");       // 국세청-대중교통 
            pGDColumn[16] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DONAT_AMT");             // 국세청-기부금  

            pGDColumn[17] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("NATIONALITY_TYPE");      // 국가타입         
            pGDColumn[18] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("REPRE_NUM");             // 주민번호 

            pGDColumn[19] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("WOMAN_YN");              // 부녀자
            pGDColumn[20] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("SINGLE_PARENT_DED_YN");  // 한부모
            pGDColumn[21] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DISABILITY_YN");         // 장애인           
            pGDColumn[22] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHILD_YN");              // 자녀양육(6세이하)

            pGDColumn[23] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_MEDIC_HIRE_INSUR_AMT");  // 기타-건강고용 보험료      
            pGDColumn[24] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_INSURE_AMT");            // 기타-보험료      
            pGDColumn[25] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_DISABILITY_INSURE_AMT"); // 기타-장애인보험료      

            pGDColumn[26] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_MEDICAL_AMT");       // 기타-의료비      
            pGDColumn[27] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_EDU_AMT");           // 기타-교육비      
            pGDColumn[28] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_CREDIT_AMT");        // 기타-신용카드    
            pGDColumn[29] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHECK_ETC_CREDIT_AMT");  // 기타-직불카드    
            pGDColumn[30] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_CASH_AMT");          // 기타-현금        
            pGDColumn[31] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_TRAD_MARKET_AMT");   // 기타-전통시장  
            pGDColumn[32] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_PUBLIC_TRANSIT_AMT");    // 기타-대중교통  
            pGDColumn[33] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("ETC_DONAT_AMT");         // 기타-기부금 
            
            //pGDColumn[34] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BASE_COUNT");            // 기본공제
            //pGDColumn[35] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("OLD_COUNT");             // 경로우대
            //pGDColumn[36] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("BIRTH_COUNT");           // 출산/입양양육
            //pGDColumn[37] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("DISABILITY_COUNT");      // 장애인
            //pGDColumn[38] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("CHILD_COUNT");           // 자녀양육(6세이하) 
            //pGDColumn[39] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("WOMAN_COUNT");           // 부녀세대 
            //pGDColumn[40] = pGrid_SUPPORT_FAMILY.GetColumnToIndex("SINGLE_PARENT_COUNT");   // 한부모 

            //---------------------------------------------------------------------------------------------------------------
            pXLColumn[0] = 4;   // 자녀양육 인원수
            pXLColumn[1] = 1;   // 관계코드         
            pXLColumn[2] = 2;   // 성명     

            pXLColumn[3] = 6;   // 기본공제         
            pXLColumn[4] = 8;   // 경로우대         
            pXLColumn[5] = 9;  // 출생입양 

            pXLColumn[6] = 12;  // 국세청-건강/고용보험 
            pXLColumn[7] = 14;  // 국세청-보험료   
            pXLColumn[8] = 17;  // 국세청-장애인보험료   
            pXLColumn[9] = 20;  // 국세청-의료비  
            pXLColumn[10] = 23;  // 국세청-교육비

            pXLColumn[11] = 26;  // 국세청-신용카드     
            pXLColumn[12] = 29; // 국세청-직불카드  
            pXLColumn[13] = 32; // 국세청-현금영수증  
            pXLColumn[14] = 35; // 국세청-전통시장이용액
            pXLColumn[15] = 38; // 국세청-대중교통이용액
            pXLColumn[16] = 41; // 국세청-기부금

            pXLColumn[17] = 1;  // 국세청-국가타입
            pXLColumn[18] = 2;  // 국세청-주민번호

            pXLColumn[19] = 6;  // 부녀자           
            pXLColumn[20] = 7;  // 한부모
            pXLColumn[21] = 8;  // 장애인
            pXLColumn[22] = 9;  // 출생입양 

            pXLColumn[23] = 12;  // 기타-건강고용등
            pXLColumn[24] = 14;  // 기타-보험료      
            pXLColumn[25] = 17;  // 기타-장애인보험료      
            pXLColumn[26] = 20;  // 기타-의료비      
            pXLColumn[27] = 23;  // 기타-교육비    

            pXLColumn[28] = 26;  // 국세청-신용카드     
            pXLColumn[29] = 29; // 국세청-직불카드  
            pXLColumn[30] = 32; // 국세청-현금영수증  
            pXLColumn[31] = 35; // 국세청-전통시장이용액
            pXLColumn[32] = 38; // 국세청-대중교통이용액
            pXLColumn[33] = 41; // 국세청-기부금 
        }

        #endregion;

        #region ----- Array Set House1 ----

        private void SetArray_House1(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[13];
            pXLColumn = new int[13];

            string vDBColumn01 = "HOUSE_LEASE_ID";
            string vDBColumn02 = "YEAR_YYYY";
            string vDBColumn03 = "PERSON_ID";
            string vDBColumn04 = "LESSOR_NAME";
            string vDBColumn05 = "LESSOR_REPRE_NUM";
            string vDBColumn06 = "LEASE_ZIP_CODE";
            string vDBColumn07 = "LEASE_ADDR1";
            string vDBColumn08 = "LEASE_ADDR2";
            string vDBColumn09 = "LEASE_TERM_FR";
            string vDBColumn10 = "LEASE_TERM_TO";
            string vDBColumn11 = "MONTLY_LEASE_AMT";
            string vDBColumn12 = "HOUSE_DED_AMT";

            pDBColumn[0] = vDBColumn01;  //HOUSE_LEASE_ID
            pDBColumn[1] = vDBColumn02;  //YEAR_YYYY
            pDBColumn[2] = vDBColumn03;  //PERSON_ID
            pDBColumn[3] = vDBColumn04;  //LESSOR_NAME
            pDBColumn[4] = vDBColumn05;  //LESSOR_REPRE_NUM
            pDBColumn[5] = vDBColumn06;  //LEASE_ZIP_CODE
            pDBColumn[6] = vDBColumn07;  //LEASE_ADDR1
            pDBColumn[7] = vDBColumn08;  //LEASE_ADDR2
            pDBColumn[8] = vDBColumn09;  //LEASE_TERM_FR
            pDBColumn[9] = vDBColumn10;  //LEASE_TERM_TO
            pDBColumn[10] = vDBColumn11;  //MONTLY_LEASE_AMT
            pDBColumn[11] = vDBColumn12;  //HOUSE_DED_AMT

            int vXLColumn01 = 0;         //HOUSE_LEASE_ID
            int vXLColumn02 = 0;         //YEAR_YYYY
            int vXLColumn03 = 0;        //PERSON_ID
            int vXLColumn04 = 1;        //LESSOR_NAME
            int vXLColumn05 = 6;        //LESSOR_REPRE_NUM
            int vXLColumn06 = 0;        //LEASE_ZIP_CODE
            int vXLColumn07 = 12;         //LEASE_ADDR1
            int vXLColumn08 = 0;         //LEASE_ADDR2
            int vXLColumn09 = 26;        //LEASE_TERM_FR
            int vXLColumn10 = 26;        //LEASE_TERM_TO
            int vXLColumn11 = 32;        //MONTLY_LEASE_AMT
            int vXLColumn12 = 38;        //HOUSE_DED_AMT


            pXLColumn[0] = vXLColumn01;  //HOUSE_LEASE_ID
            pXLColumn[1] = vXLColumn02;  //YEAR_YYYY
            pXLColumn[2] = vXLColumn03;  //PERSON_ID
            pXLColumn[3] = vXLColumn04;  //LESSOR_NAME
            pXLColumn[4] = vXLColumn05;  //LESSOR_REPRE_NUM
            pXLColumn[5] = vXLColumn06;  //LEASE_ZIP_CODE
            pXLColumn[6] = vXLColumn07;  //LEASE_ADDR1
            pXLColumn[7] = vXLColumn08;  //LEASE_ADDR2
            pXLColumn[8] = vXLColumn09;  //LEASE_TERM_FR
            pXLColumn[9] = vXLColumn10;  //LEASE_TERM_TO
            pXLColumn[10] = vXLColumn11;  //MONTLY_LEASE_AMT
            pXLColumn[11] = vXLColumn12;  //HOUSE_DED_AMT

        }

        #endregion;

        #region ----- Array Set House2 ----

        private void SetArray_House2(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[21];
            pXLColumn = new int[21];

            string vDBColumn01 = "HOUSE_LEASE_ID";
            string vDBColumn02 = "YEAR_YYYY";
            string vDBColumn03 = "PERSON_ID";
            string vDBColumn04 = "LOANER_NAME";
            string vDBColumn05 = "LOANER_REPRE_NUM";
            string vDBColumn06 = "LOAN_TERM_FR";
            string vDBColumn07 = "LOAN_TERM_TO";
            string vDBColumn08 = "LOAN_INTEREST_RATE";
            string vDBColumn09 = "LOAN_TOT_AMT";
            string vDBColumn10 = "LOAN_AMT";
            string vDBColumn11 = "LOAN_INTEREST_AMT";
            string vDBColumn12 = "HOUSE_DED_AMT";
            string vDBColumn13 = "LESSOR_NAME";
            string vDBColumn14 = "LESSOR_REPRE_NUM";
            string vDBColumn15 = "LEASE_ZIP_CODE";
            string vDBColumn16 = "LEASE_ADDR1";
            string vDBColumn17 = "LEASE_ADDR2";
            string vDBColumn18 = "LEASE_TERM_FR";
            string vDBColumn19 = "LEASE_TERM_TO";
            string vDBColumn20 = "DEPOSIT_AMT";

            pDBColumn[0] = vDBColumn01;  //HOUSE_LEASE_ID
            pDBColumn[1] = vDBColumn02;  //YEAR_YYYY
            pDBColumn[2] = vDBColumn03;  //PERSON_ID
            pDBColumn[3] = vDBColumn04;  //LOANER_NAME
            pDBColumn[4] = vDBColumn05;  //LOANER_REPRE_NUM
            pDBColumn[5] = vDBColumn06;  //LOAN_TERM_FR
            pDBColumn[6] = vDBColumn07;  //LOAN_TERM_TO
            pDBColumn[7] = vDBColumn08;  //LOAN_INTEREST_RATE
            pDBColumn[8] = vDBColumn09;  //LOAN_TOT_AMT
            pDBColumn[9] = vDBColumn10;  //LOAN_AMT
            pDBColumn[10] = vDBColumn11;  //LOAN_INTEREST_AMT
            pDBColumn[11] = vDBColumn12;  //HOUSE_DED_AMT
            pDBColumn[12] = vDBColumn13;  //LESSOR_NAME
            pDBColumn[13] = vDBColumn14;  //LESSOR_REPRE_NUM
            pDBColumn[14] = vDBColumn15;  //LEASE_ZIP_CODE
            pDBColumn[15] = vDBColumn16;  //LEASE_ADDR1
            pDBColumn[16] = vDBColumn17;  //LEASE_ADDR2
            pDBColumn[17] = vDBColumn18;  //LEASE_TERM_FR
            pDBColumn[18] = vDBColumn19;  //LEASE_TERM_TO
            pDBColumn[19] = vDBColumn20;  //DEPOSIT_AMT


            int vXLColumn01 = 0;         //HOUSE_LEASE_ID
            int vXLColumn02 = 0;         //YEAR_YYYY
            int vXLColumn03 = 0;        //PERSON_ID
            int vXLColumn04 = 1;        //LOANER_NAME
            int vXLColumn05 = 6;        //LOANER_REPRE_NUM
            int vXLColumn06 = 12;        //LOAN_TERM_FR
            int vXLColumn07 = 12;         //LOAN_TERM_TO
            int vXLColumn08 = 18;         //LOAN_INTEREST_RATE
            int vXLColumn09 = 22;        //LOAN_TOT_AMT
            int vXLColumn10 = 28;        //LOAN_AMT
            int vXLColumn11 = 33;        //LOAN_INTEREST_AMT
            int vXLColumn12 = 38;        //HOUSE_DED_AMT
            int vXLColumn13 = 1;        //LESSOR_NAME
            int vXLColumn14 = 6;        //LESSOR_REPRE_NUM
            int vXLColumn15 = 0;        //LEASE_ZIP_CODE
            int vXLColumn16 = 12;        //LEASE_ADDR1
            int vXLColumn17 = 0;        //LEASE_ADDR2
            int vXLColumn18 = 26;        //LEASE_TERM_FR
            int vXLColumn19 = 0;        //LEASE_TERM_TO
            int vXLColumn20 = 35;        //DEPOSIT_AMT


            pXLColumn[0] = vXLColumn01;  //HOUSE_LEASE_ID
            pXLColumn[1] = vXLColumn02;  //YEAR_YYYY
            pXLColumn[2] = vXLColumn03;  //PERSON_ID
            pXLColumn[3] = vXLColumn04;  //LOANER_NAME
            pXLColumn[4] = vXLColumn05;  //LOANER_REPRE_NUM
            pXLColumn[5] = vXLColumn06;  //LOAN_TERM_FR
            pXLColumn[6] = vXLColumn07;  //LOAN_TERM_TO
            pXLColumn[7] = vXLColumn08;  //LOAN_INTEREST_RATE
            pXLColumn[8] = vXLColumn09;  //LOAN_TOT_AMT
            pXLColumn[9] = vXLColumn10;  //LOAN_AMT
            pXLColumn[10] = vXLColumn11;  //LOAN_INTEREST_AMT
            pXLColumn[11] = vXLColumn12;  //HOUSE_DED_AMT
            pXLColumn[12] = vXLColumn13;  //LESSOR_NAME
            pXLColumn[13] = vXLColumn14;  //LESSOR_REPRE_NUM
            pXLColumn[14] = vXLColumn15;  //LEASE_ZIP_CODE
            pXLColumn[15] = vXLColumn16;  //LEASE_ADDR1
            pXLColumn[16] = vXLColumn17;  //LEASE_ADDR2
            pXLColumn[17] = vXLColumn18;  //LEASE_TERM_FR
            pXLColumn[18] = vXLColumn19;  //LEASE_TERM_TO
            pXLColumn[19] = vXLColumn20;  //DEPOSIT_AMT

        }

        #endregion;

        #region ----- Array Set House3 ----

        private void SetArray_House3(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[15];
            pXLColumn = new int[15];

            string vDBColumn01 = "HOUSE_LEASE_ID";
            string vDBColumn02 = "YEAR_YYYY";
            string vDBColumn03 = "PERSON_ID";
            string vDBColumn04 = "LESSOR_NAME";
            string vDBColumn05 = "LESSOR_REPRE_NUM";
            string vDBColumn06 = "LEASE_ZIP_CODE";
            string vDBColumn07 = "LEASE_ADDR1";
            string vDBColumn08 = "LEASE_ADDR2";
            string vDBColumn09 = "LEASE_TERM_FR";
            string vDBColumn10 = "LEASE_TERM_TO";
            string vDBColumn11 = "MONTLY_LEASE_AMT";
            string vDBColumn12 = "HOUSE_DED_AMT";
            string vDBColumn13 = "HOUSE_TYPE_NAME";
            string vDBColumn14 = "HOUSE_AREA";


            pDBColumn[0] = vDBColumn01;  //HOUSE_LEASE_ID
            pDBColumn[1] = vDBColumn02;  //YEAR_YYYY
            pDBColumn[2] = vDBColumn03;  //PERSON_ID
            pDBColumn[3] = vDBColumn04;  //LESSOR_NAME
            pDBColumn[4] = vDBColumn05;  //LESSOR_REPRE_NUM
            pDBColumn[5] = vDBColumn06;  //LEASE_ZIP_CODE
            pDBColumn[6] = vDBColumn07;  //LEASE_ADDR1
            pDBColumn[7] = vDBColumn08;  //LEASE_ADDR2
            pDBColumn[8] = vDBColumn09;  //LEASE_TERM_FR
            pDBColumn[9] = vDBColumn10;  //LEASE_TERM_TO
            pDBColumn[10] = vDBColumn11;  //MONTLY_LEASE_AMT
            pDBColumn[11] = vDBColumn12;  //HOUSE_DED_AMT
            pDBColumn[12] = vDBColumn13;  //HOUSE_TYPE_NAME
            pDBColumn[13] = vDBColumn14;  //HOUSE_AREA


            int vXLColumn01 = 0;         //HOUSE_LEASE_ID
            int vXLColumn02 = 0;         //YEAR_YYYY
            int vXLColumn03 = 0;        //PERSON_ID
            int vXLColumn04 = 1;        //LESSOR_NAME
            int vXLColumn05 = 6;        //LESSOR_REPRE_NUM
            int vXLColumn06 = 0;        //LEASE_ZIP_CODE
            int vXLColumn07 = 19;         //LEASE_ADDR1
            int vXLColumn08 = 0;         //LEASE_ADDR2
            int vXLColumn09 = 26;        //LEASE_TERM_FR
            int vXLColumn10 = 30;        //LEASE_TERM_TO
            int vXLColumn11 = 34;        //MONTLY_LEASE_AMT
            int vXLColumn12 = 39;        //HOUSE_DED_AMT
            int vXLColumn13 = 12;        //HOUSE_TYPE_NAME
            int vXLColumn14 = 15;        //HOUSE_AREA


            pXLColumn[0] = vXLColumn01;  //HOUSE_LEASE_ID
            pXLColumn[1] = vXLColumn02;  //YEAR_YYYY
            pXLColumn[2] = vXLColumn03;  //PERSON_ID
            pXLColumn[3] = vXLColumn04;  //LESSOR_NAME
            pXLColumn[4] = vXLColumn05;  //LESSOR_REPRE_NUM
            pXLColumn[5] = vXLColumn06;  //LEASE_ZIP_CODE
            pXLColumn[6] = vXLColumn07;  //LEASE_ADDR1
            pXLColumn[7] = vXLColumn08;  //LEASE_ADDR2
            pXLColumn[8] = vXLColumn09;  //LEASE_TERM_FR
            pXLColumn[9] = vXLColumn10;  //LEASE_TERM_TO
            pXLColumn[10] = vXLColumn11;  //MONTLY_LEASE_AMT
            pXLColumn[11] = vXLColumn12;  //HOUSE_DED_AMT
            pXLColumn[12] = vXLColumn13;  //HOUSE_TYPE_NAME
            pXLColumn[13] = vXLColumn14;  //HOUSE_AREA


        }

        #endregion;

        #region ----- Array Set House4 ----

        private void SetArray_House4(out string[] pDBColumn, out int[] pXLColumn)
        {
            pDBColumn = new string[23];
            pXLColumn = new int[23];

            string vDBColumn01 = "HOUSE_LEASE_ID";
            string vDBColumn02 = "YEAR_YYYY";
            string vDBColumn03 = "PERSON_ID";
            string vDBColumn04 = "LOANER_NAME";
            string vDBColumn05 = "LOANER_REPRE_NUM";
            string vDBColumn06 = "LOAN_TERM_FR";
            string vDBColumn07 = "LOAN_TERM_TO";
            string vDBColumn08 = "LOAN_INTEREST_RATE";
            string vDBColumn09 = "LOAN_TOT_AMT";
            string vDBColumn10 = "LOAN_AMT";
            string vDBColumn11 = "LOAN_INTEREST_AMT";
            string vDBColumn12 = "HOUSE_DED_AMT";
            string vDBColumn13 = "LESSOR_NAME";
            string vDBColumn14 = "LESSOR_REPRE_NUM";
            string vDBColumn15 = "LEASE_ZIP_CODE";
            string vDBColumn16 = "LEASE_ADDR1";
            string vDBColumn17 = "LEASE_ADDR2";
            string vDBColumn18 = "LEASE_TERM_FR";
            string vDBColumn19 = "LEASE_TERM_TO";
            string vDBColumn20 = "DEPOSIT_AMT";
            string vDBColumn21 = "HOUSE_TYPE_NAME";
            string vDBColumn22 = "HOUSE_AREA";


            pDBColumn[0] = vDBColumn01;  //HOUSE_LEASE_ID
            pDBColumn[1] = vDBColumn02;  //YEAR_YYYY
            pDBColumn[2] = vDBColumn03;  //PERSON_ID
            pDBColumn[3] = vDBColumn04;  //LOANER_NAME
            pDBColumn[4] = vDBColumn05;  //LOANER_REPRE_NUM
            pDBColumn[5] = vDBColumn06;  //LOAN_TERM_FR
            pDBColumn[6] = vDBColumn07;  //LOAN_TERM_TO
            pDBColumn[7] = vDBColumn08;  //LOAN_INTEREST_RATE
            pDBColumn[8] = vDBColumn09;  //LOAN_TOT_AMT
            pDBColumn[9] = vDBColumn10;  //LOAN_AMT
            pDBColumn[10] = vDBColumn11;  //LOAN_INTEREST_AMT
            pDBColumn[11] = vDBColumn12;  //HOUSE_DED_AMT
            pDBColumn[12] = vDBColumn13;  //LESSOR_NAME
            pDBColumn[13] = vDBColumn14;  //LESSOR_REPRE_NUM
            pDBColumn[14] = vDBColumn15;  //LEASE_ZIP_CODE
            pDBColumn[15] = vDBColumn16;  //LEASE_ADDR1
            pDBColumn[16] = vDBColumn17;  //LEASE_ADDR2
            pDBColumn[17] = vDBColumn18;  //LEASE_TERM_FR
            pDBColumn[18] = vDBColumn19;  //LEASE_TERM_TO
            pDBColumn[19] = vDBColumn20;  //DEPOSIT_AMT
            pDBColumn[20] = vDBColumn21;  //HOUSE_TYPE_NAME
            pDBColumn[21] = vDBColumn22;  //HOUSE_AREA


            int vXLColumn01 = 0;         //HOUSE_LEASE_ID
            int vXLColumn02 = 0;         //YEAR_YYYY
            int vXLColumn03 = 0;        //PERSON_ID
            int vXLColumn04 = 1;        //LOANER_NAME
            int vXLColumn05 = 6;        //LOANER_REPRE_NUM
            int vXLColumn06 = 12;        //LOAN_TERM_FR
            int vXLColumn07 = 12;         //LOAN_TERM_TO
            int vXLColumn08 = 18;         //LOAN_INTEREST_RATE
            int vXLColumn09 = 22;        //LOAN_TOT_AMT
            int vXLColumn10 = 28;        //LOAN_AMT
            int vXLColumn11 = 33;        //LOAN_INTEREST_AMT
            int vXLColumn12 = 38;        //HOUSE_DED_AMT
            int vXLColumn13 = 1;        //LESSOR_NAME
            int vXLColumn14 = 6;        //LESSOR_REPRE_NUM
            int vXLColumn15 = 0;        //LEASE_ZIP_CODE
            int vXLColumn16 = 19;        //LEASE_ADDR1
            int vXLColumn17 = 0;        //LEASE_ADDR2
            int vXLColumn18 = 26;        //LEASE_TERM_FR
            int vXLColumn19 = 31;        //LEASE_TERM_TO
            int vXLColumn20 = 36;        //DEPOSIT_AMT
            int vXLColumn21 = 12;        //HOUSE_TYPE_NAME
            int vXLColumn22 = 15;        //HOUSE_AREA


            pXLColumn[0] = vXLColumn01;  //HOUSE_LEASE_ID
            pXLColumn[1] = vXLColumn02;  //YEAR_YYYY
            pXLColumn[2] = vXLColumn03;  //PERSON_ID
            pXLColumn[3] = vXLColumn04;  //LOANER_NAME
            pXLColumn[4] = vXLColumn05;  //LOANER_REPRE_NUM
            pXLColumn[5] = vXLColumn06;  //LOAN_TERM_FR
            pXLColumn[6] = vXLColumn07;  //LOAN_TERM_TO
            pXLColumn[7] = vXLColumn08;  //LOAN_INTEREST_RATE
            pXLColumn[8] = vXLColumn09;  //LOAN_TOT_AMT
            pXLColumn[9] = vXLColumn10;  //LOAN_AMT
            pXLColumn[10] = vXLColumn11;  //LOAN_INTEREST_AMT
            pXLColumn[11] = vXLColumn12;  //HOUSE_DED_AMT
            pXLColumn[12] = vXLColumn13;  //LESSOR_NAME
            pXLColumn[13] = vXLColumn14;  //LESSOR_REPRE_NUM
            pXLColumn[14] = vXLColumn15;  //LEASE_ZIP_CODE
            pXLColumn[15] = vXLColumn16;  //LEASE_ADDR1
            pXLColumn[16] = vXLColumn17;  //LEASE_ADDR2
            pXLColumn[17] = vXLColumn18;  //LEASE_TERM_FR
            pXLColumn[18] = vXLColumn19;  //LEASE_TERM_TO
            pXLColumn[19] = vXLColumn20;  //DEPOSIT_AMT
            pXLColumn[20] = vXLColumn21;  //HOUSE_TYPE_NAME
            pXLColumn[21] = vXLColumn22;  //HOUSE_AREA

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

        #region ----- Send ORG ----

        private void SendORG()
        {
            //mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        }

        #endregion;

        #region ----- XLLINE14 -----

        private int XLLine14(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, object pPrintDate, string pPrint_Type, object pPrint_Type_Desc)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            //System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //----[ 1 page ]------------------------------------------------------------------------------------------------------

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 거주 구분(거주자1/거주자2)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //거주자1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //거주자 2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //거주자1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //거주자 2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 거주지국
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 거주지국코드
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 내외국인 구분(내국인1/외국인9)
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //내국인1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //외국인9이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //내국인1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //외국인9이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 외국인단일세율적용
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "Y") //여1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //부2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 3), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //여1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //부2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 3), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 출력 용도 구분
                vXLColumnIndex = 13;
                vObject = pPrint_Type_Desc;
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

                // 국적
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 국적코드
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 세대주 구분(세대주1/세대원2)
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //세대주1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //세대원2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //세대주1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //세대원2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 연말정산구분
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "계속근로") //계속근로1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //중도퇴사2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //계속근로1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //중도퇴사2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 법인명(상호)
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 대표자(성명)    
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업자등록번호
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 소재지(주소)
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 성명
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                string sName = vConvertString;

                // 주민번호
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                string sPersonNumber = vConvertString;

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주소
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 종전
                vGDColumnIndex = pGDColumn[164];
                vXLColumnIndex = pXLColumn[164];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)근무처명
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1근무처명
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2근무처명
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3근무처명
                vGDColumnIndex = pGDColumn[165];
                vXLColumnIndex = pXLColumn[165];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)사업자번호
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1사업잡번호
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
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

                // 종(전)2사업잡번호
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3사업잡번호
                vGDColumnIndex = pGDColumn[166];
                vXLColumnIndex = pXLColumn[166];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)근무기간
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1근무기간
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2근무기간
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3근무기간
                vGDColumnIndex = pGDColumn[167];
                vXLColumnIndex = pXLColumn[167];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)감면기간
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1감면기간
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2감면기간
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3감면기간
                vGDColumnIndex = pGDColumn[168];
                vXLColumnIndex = pXLColumn[168];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)급여 
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1급여
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2급여
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3급여
                vGDColumnIndex = pGDColumn[169];
                vXLColumnIndex = pXLColumn[169];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)상여
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1상여
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2상여
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3상여
                vGDColumnIndex = pGDColumn[170];
                vXLColumnIndex = pXLColumn[170];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)인정상여 
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1인정상여
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2인정상여
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3인정상여
                vGDColumnIndex = pGDColumn[171];
                vXLColumnIndex = pXLColumn[171];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)주식매수선택권
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1주식매수선택권
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2주식매수선택권
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3주식매수선택권
                vGDColumnIndex = pGDColumn[172];
                vXLColumnIndex = pXLColumn[172];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)우리사주조합인출금
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1우리사주조합인출금
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2우리사주조합인출금
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3우리사주조합인출금
                vGDColumnIndex = pGDColumn[173];
                vXLColumnIndex = pXLColumn[173];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)임원퇴직소득금액 한도초과액
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1임원퇴직소득금액 한도초과액
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2임원퇴직소득금액 한도초과액
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3임원퇴직소득금액 한도초과액
                vGDColumnIndex = pGDColumn[174];
                vXLColumnIndex = pXLColumn[174];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주(현)계
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1계
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2계
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3계
                vGDColumnIndex = pGDColumn[175];
                vXLColumnIndex = pXLColumn[175];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //--------------------------------------------------------------------------------------------------------------------
                // 비과세 및 감면 소득 명세
                //--------------------------------------------------------------------------------------------------------------------

                // 비과세_주(현)국외근로
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1국외근로
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2국외근로
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3국외근로
                vGDColumnIndex = pGDColumn[176];
                vXLColumnIndex = pXLColumn[176];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)야간근로수당
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1야간근로수당
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2야간근로수당
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3야간근로수당
                vGDColumnIndex = pGDColumn[177];
                vXLColumnIndex = pXLColumn[177];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)출산/보육수당
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1출산/보육수당
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2출산/보육수당
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2출산/보육수당
                vGDColumnIndex = pGDColumn[178];
                vXLColumnIndex = pXLColumn[178];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)연구보조비
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1연구보조비
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2연구보조비
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2연구보조비
                vGDColumnIndex = pGDColumn[179];
                vXLColumnIndex = pXLColumn[179];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // 중소기업에 취업하는 청년에 대한 소득세 감면 존재
                vGDColumnIndex = pGDColumn[187];
                vXLColumnIndex = pXLColumn[187];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }

                // 중소기업에 취업하는 청년에 대한 소득세 감면1
                vGDColumnIndex = pGDColumn[188];
                vXLColumnIndex = pXLColumn[188];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }


                // 중소기업에 취업하는 청년에 대한 소득세 감면2
                vGDColumnIndex = pGDColumn[189];
                vXLColumnIndex = pXLColumn[189];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }

                // 중소기업에 취업하는 청년에 대한 소득세 감면3
                vGDColumnIndex = pGDColumn[190];
                vXLColumnIndex = pXLColumn[190];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 비과세_주(현)수련보조수당
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1수련보조수당
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2수련보조수당
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3수련보조수당
                vGDColumnIndex = pGDColumn[180];
                vXLColumnIndex = pXLColumn[180];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)비과세소득 계
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1비과세소득 계
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2비과세소득 계
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3비과세소득 계
                vGDColumnIndex = pGDColumn[181];
                vXLColumnIndex = pXLColumn[181];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)감면소득계
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1감면소득계
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2감면소득계
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3감면소득계
                vGDColumnIndex = pGDColumn[182];
                vXLColumnIndex = pXLColumn[182];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                //--------------------------------------------------------------------------------------------------------------------
                // 세액 명세
                //--------------------------------------------------------------------------------------------------------------------

                // 결정세액_소득세
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_지방소득세   
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_농특세
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)1사업자번호 
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1소득세
                vGDColumnIndex = pGDColumn[74];
                vXLColumnIndex = pXLColumn[74];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1지방소득세
                vGDColumnIndex = pGDColumn[75];
                vXLColumnIndex = pXLColumn[75];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1농특세
                vGDColumnIndex = pGDColumn[76];
                vXLColumnIndex = pXLColumn[76];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)2사업자번호 
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2소득세
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2지방소득세
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2농특세
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)3사업자번호 
                vGDColumnIndex = pGDColumn[183];
                vXLColumnIndex = pXLColumn[183];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)3소득세
                vGDColumnIndex = pGDColumn[184];
                vXLColumnIndex = pXLColumn[184];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)3지방소득세
                vGDColumnIndex = pGDColumn[185];
                vXLColumnIndex = pXLColumn[185];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)3농특세
                vGDColumnIndex = pGDColumn[186];
                vXLColumnIndex = pXLColumn[186];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_주(현)소득세 
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_주(현)지방소득세
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //기납부세액_주(현)농특세
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_소득세 
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_지방소득세
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_농특세
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[86];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 5;
                //-------------------------------------------------------------------

                // 날짜
                vXLColumnIndex = 28;
                vObject = pPrintDate;
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

                // 징수의무자
                vXLColumnIndex = 23;
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WITHHOLDING_OWNER");
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 받는자  
                if (pPrint_Type == "1")
                {//소득자 보관용
                    vGDColumnIndex = pGDColumn[231];
                    vXLColumnIndex = pXLColumn[231];
                    vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                }
                else if (pPrint_Type == "2")
                {
                    vGDColumnIndex = pGDColumn[232];
                    vXLColumnIndex = pXLColumn[232];
                    vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                }
                else
                {
                    vGDColumnIndex = pGDColumn[233];
                    vXLColumnIndex = pXLColumn[233];
                    vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                } 
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
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                //----[ 2 page ]------------------------------------------------------------------------------------------------------

                // 2page 상단에 소득자 성명 및 주민번호 출력 표시되는 부분
                string sPrintPersinInfo = sName + "(" + sPersonNumber + ")";
                mPrinting.XLSetCell(vXLine, 24, sPrintPersinInfo);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //21. 총급여
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[87];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //50. 종합소득과세표준
                vGDColumnIndex = pGDColumn[149];
                vXLColumnIndex = pXLColumn[149];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                //22. 근로소득공제
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[88];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 51. 산출세액
                vGDColumnIndex = pGDColumn[150];
                vXLColumnIndex = pXLColumn[150];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //23. 근로소득금액
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[89];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //24. 기본(본인)
                vGDColumnIndex = pGDColumn[90];
                vXLColumnIndex = pXLColumn[90];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //52. 소득세법
                vGDColumnIndex = pGDColumn[151];
                vXLColumnIndex = pXLColumn[151];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //25. 기본(배우자)
                vGDColumnIndex = pGDColumn[91];
                vXLColumnIndex = pXLColumn[91];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //26. 기본(부양인원 - 인원)  
                vGDColumnIndex = pGDColumn[92];
                vXLColumnIndex = pXLColumn[92];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //26. 기본(부양인원 - 금액) 
                vGDColumnIndex = pGDColumn[93];
                vXLColumnIndex = pXLColumn[93];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //53. 「조세특례제한법」(<54>-1제외)
                vGDColumnIndex = pGDColumn[152];
                vXLColumnIndex = pXLColumn[152];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                //27. 추가공제(경로수 - 인원)
                vGDColumnIndex = pGDColumn[94];
                vXLColumnIndex = pXLColumn[94];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //27. 추가공제(경로수 - 금액)
                vGDColumnIndex = pGDColumn[95];
                vXLColumnIndex = pXLColumn[95];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //54. 「조세특례제한법」 제30조
                vGDColumnIndex = pGDColumn[153];
                vXLColumnIndex = pXLColumn[153];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //28. 추가공제(장애인 - 인원)
                vGDColumnIndex = pGDColumn[96];
                vXLColumnIndex = pXLColumn[96];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //28. 추가공제(장애인 - 금액)
                vGDColumnIndex = pGDColumn[97];
                vXLColumnIndex = pXLColumn[97];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //29. 추가공제(부녀세대)
                vGDColumnIndex = pGDColumn[98];
                vXLColumnIndex = pXLColumn[98];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //55. 조세조약
                vGDColumnIndex = pGDColumn[154];
                vXLColumnIndex = pXLColumn[154];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //30. 추가공제(한부모가족 - 금액)
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //31. 국민연금보험료공제
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //56. 세 액 감 면 계
                vGDColumnIndex = pGDColumn[155];
                vXLColumnIndex = pXLColumn[155];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;  //76
                //-------------------------------------------------------------------
                //32.가. 공무원연금
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //32.나. 군인연금
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //57. 근로소득
                vGDColumnIndex = pGDColumn[156];
                vXLColumnIndex = pXLColumn[156];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;  //78
                //-------------------------------------------------------------------
                //32.다 사립합교교직원연금
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;  //79 
                //-------------------------------------------------------------------
                //라. 별정우체국연금
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                
                //58. 2014-세공(자녀양육) 인원
                vGDColumnIndex = pGDColumn[196];
                vXLColumnIndex = pXLColumn[196];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //58. 2014-세공(자녀양육) 금액
                vGDColumnIndex = pGDColumn[197];
                vXLColumnIndex = pXLColumn[197];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;  //80 
                //-------------------------------------------------------------------
                //58. 2014-세공(자녀양육 6세이하) 인원
                vGDColumnIndex = pGDColumn[198];
                vXLColumnIndex = pXLColumn[198];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //58. 2014-세공(자녀양육 6세이하) 금액
                vGDColumnIndex = pGDColumn[199];
                vXLColumnIndex = pXLColumn[199];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;  //81
                //-------------------------------------------------------------------

                //33.가 건강보험료(노인장기요양보험료 포함)
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //58. 2014-세공(자녀양육 출생입양) 인원
                vGDColumnIndex = pGDColumn[200];
                vXLColumnIndex = pXLColumn[200];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //58. 2014-세공(자녀양육 6출생입양) 금액
                vGDColumnIndex = pGDColumn[201];
                vXLColumnIndex = pXLColumn[201];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;  //82 
                //-------------------------------------------------------------------
                //59. 2014-과학기술인 공제금액
                vGDColumnIndex = pGDColumn[203];
                vXLColumnIndex = pXLColumn[203];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //83
                //-------------------------------------------------------------------

                //33.나 고용보험료
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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


                //59. 2014-과학기술인 세액공제
                vGDColumnIndex = pGDColumn[204];
                vXLColumnIndex = pXLColumn[204];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //84
                //-------------------------------------------------------------------

                //37.가 주택임차차입금원리금상환액-대출기관
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //60. 2014-퇴직연금  공제금액
                vGDColumnIndex = pGDColumn[205];
                vXLColumnIndex = pXLColumn[205];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //85
                //-------------------------------------------------------------------
                
                //60. 2014-퇴직연금  세액공제
                vGDColumnIndex = pGDColumn[206];
                vXLColumnIndex = pXLColumn[206];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //86
                //-------------------------------------------------------------------

                //37.가 주택임차차입금원리금상환액-거주자
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                //61. 2014-연금저축 공제금액
                vGDColumnIndex = pGDColumn[207];
                vXLColumnIndex = pXLColumn[207];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;  //87
                //-------------------------------------------------------------------
                //61.  2014-연금저축 세액공제
                vGDColumnIndex = pGDColumn[208];
                vXLColumnIndex = pXLColumn[208];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;  //88
                //-------------------------------------------------------------------

                //37.나 장기주택저당차입금이자상환액 - 2011이전 (15년미만)
                vGDColumnIndex = pGDColumn[125];
                vXLColumnIndex = pXLColumn[125];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //62. 2014-세감 ( 보장성보험 공제금액)
                vGDColumnIndex = pGDColumn[209];
                vXLColumnIndex = pXLColumn[209];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //89
                //-------------------------------------------------------------------

                //37.나 장기주택저당차입금이자상환액 - 2011이전 (15년~29년)
                vGDColumnIndex = pGDColumn[126];
                vXLColumnIndex = pXLColumn[126];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                
                //62. 2014-세감 ( 보장성보험 세액공제)
                vGDColumnIndex = pGDColumn[210];
                vXLColumnIndex = pXLColumn[210];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //90
                //-------------------------------------------------------------------

                //37.나 장기주택저당차입금이자상환액 - 2011이전 (30년 이상)
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //62. 2014-세감 ( 장애인보장성보험 공제금액)
                vGDColumnIndex = pGDColumn[211];
                vXLColumnIndex = pXLColumn[211];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //91
                //-------------------------------------------------------------------

                // 장기주택저당차입금이자상환액 - 2012이후 (고정금리비거치 상환대출)
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //62. 2014-세감 ( 장애인보장성보험 공제금액)
                vGDColumnIndex = pGDColumn[212];
                vXLColumnIndex = pXLColumn[212];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //92
                //-------------------------------------------------------------------

                // 2014-세감 (의료비 공제금액)
                vGDColumnIndex = pGDColumn[213];
                vXLColumnIndex = pXLColumn[213];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //93
                //-------------------------------------------------------------------

                // 2014-세감 (의료비 세액공제)
                vGDColumnIndex = pGDColumn[214];
                vXLColumnIndex = pXLColumn[214];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //94
                //-------------------------------------------------------------------
                 
                // 장기주택저당차입금이자상환액 - 2012이후 (기타)
                vGDColumnIndex = pGDColumn[129];
                vXLColumnIndex = pXLColumn[129];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 2014-세감 (교육비 공제금액)
                vGDColumnIndex = pGDColumn[215];
                vXLColumnIndex = pXLColumn[215];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //95
                //-------------------------------------------------------------------

                // 2014-세감 (교육비 세액공제) 
                vGDColumnIndex = pGDColumn[216];
                vXLColumnIndex = pXLColumn[216];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //96
                //-------------------------------------------------------------------

                // 2014-특별소득공제(기부금이월분)
                vGDColumnIndex = pGDColumn[191];
                vXLColumnIndex = pXLColumn[191];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 2014-세감 (정치자금기부금-10만원이하) 공제금액
                vGDColumnIndex = pGDColumn[217];
                vXLColumnIndex = pXLColumn[217];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //97
                //-------------------------------------------------------------------

                // 2014-특별소득공제 합계
                vGDColumnIndex = pGDColumn[226];
                vXLColumnIndex = pXLColumn[226];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 2014-세감 (정치자금기부금-10만원이하) 세액공제
                vGDColumnIndex = pGDColumn[218];
                vXLColumnIndex = pXLColumn[218];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //98
                //-------------------------------------------------------------------

                // 차감소득금액
                vGDColumnIndex = pGDColumn[136];
                vXLColumnIndex = pXLColumn[136];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
 
                // 2014-세감 (정치자금기부금-10만원초과) 공제금액
                vGDColumnIndex = pGDColumn[212];
                vXLColumnIndex = pXLColumn[212];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //99
                //-------------------------------------------------------------------

                // 2014-세감 (정치자금기부금-10만원초과) 세액공제
                vGDColumnIndex = pGDColumn[220];
                vXLColumnIndex = pXLColumn[220];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //100
                //-------------------------------------------------------------------

                // 개인연금저축소득공제
                vGDColumnIndex = pGDColumn[137];
                vXLColumnIndex = pXLColumn[137];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 2014-세감 (법정기부금) 공제금액
                vGDColumnIndex = pGDColumn[221];
                vXLColumnIndex = pXLColumn[221];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //101
                //-------------------------------------------------------------------                               

                // 2014-세감 (법정기부금) 세액공제
                vGDColumnIndex = pGDColumn[222];
                vXLColumnIndex = pXLColumn[222];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //102
                //-------------------------------------------------------------------
               
                // 소기업/소상공인 공제부금 소득공제
                vGDColumnIndex = pGDColumn[138];
                vXLColumnIndex = pXLColumn[138];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                
                //우리사주조합 기부금 2015년도분 공제대상 -> 현재 없음 추후 추가 해야 함
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //103
                //-------------------------------------------------------------------
                //우리사주조합 기부금 2015년도분 공제세액 -> 현재 없음 추후 추가 해야 함

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;    //104
                //-------------------------------------------------------------------

                // 청약저축
                vGDColumnIndex = pGDColumn[139];
                vXLColumnIndex = pXLColumn[139];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 2014-세감 (지정기부금) 공제금액
                vGDColumnIndex = pGDColumn[223];
                vXLColumnIndex = pXLColumn[223];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //105
                //-------------------------------------------------------------------

                // 주택청약종합저축
                vGDColumnIndex = pGDColumn[140];
                vXLColumnIndex = pXLColumn[140];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 2014-세감 (지정기부금) 세액공제
                vGDColumnIndex = pGDColumn[224];
                vXLColumnIndex = pXLColumn[224];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //106
                //-------------------------------------------------------------------

                // 근로자주택마련저축
                vGDColumnIndex = pGDColumn[141];
                vXLColumnIndex = pXLColumn[141];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                                
                // 2014-특별세액공제 계   
                vGDColumnIndex = pGDColumn[227];
                vXLColumnIndex = pXLColumn[220];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //107
                //-------------------------------------------------------------------
                // 투자조합출자등소득공제
                vGDColumnIndex = pGDColumn[142];
                vXLColumnIndex = pXLColumn[142];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 표준공제
                vGDColumnIndex = pGDColumn[135];
                vXLColumnIndex = pXLColumn[135];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //108
                //-------------------------------------------------------------------
                                
                // 납세조합공제
                vGDColumnIndex = pGDColumn[157];
                vXLColumnIndex = pXLColumn[157];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //109
                //-------------------------------------------------------------------

                // 신용카드등소득공제
                vGDColumnIndex = pGDColumn[143];
                vXLColumnIndex = pXLColumn[143];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //110
                //-------------------------------------------------------------------

                // 주택차입금
                vGDColumnIndex = pGDColumn[158];
                vXLColumnIndex = pXLColumn[158];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //111 
                //-------------------------------------------------------------------

                // 우리사주조합소득공제
                vGDColumnIndex = pGDColumn[144];
                vXLColumnIndex = pXLColumn[144];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 2;    //113
                //-------------------------------------------------------------------

                // 2014-그밖의소득공제(우리사주조합기부금)
                vGDColumnIndex = pGDColumn[192];
                vXLColumnIndex = pXLColumn[192];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 외국납부
                vGDColumnIndex = pGDColumn[160];
                vXLColumnIndex = pXLColumn[160];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 2;    //115
                //-------------------------------------------------------------------

                //  고용유지중소기업근로자
                vGDColumnIndex = pGDColumn[145];
                vXLColumnIndex = pXLColumn[145];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //116
                //-------------------------------------------------------------------

                //  월세액 - 공제대상금액
                vGDColumnIndex = pGDColumn[229];
                vXLColumnIndex = pXLColumn[229];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //117
                //-------------------------------------------------------------------

                // 목돈 안드는 전세이자상환액 
                vGDColumnIndex = pGDColumn[146];
                vXLColumnIndex = pXLColumn[146];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //  월세액 - 세액공제액
                vGDColumnIndex = pGDColumn[230];
                vXLColumnIndex = pXLColumn[230];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //118
                //-------------------------------------------------------------------

                // 2014-세액공제 계 
                vGDColumnIndex = pGDColumn[228];
                vXLColumnIndex = pXLColumn[228];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //119
                //-------------------------------------------------------------------

                // 2014- 장기집합투자증권저축
                vGDColumnIndex = pGDColumn[194];
                vXLColumnIndex = pXLColumn[194];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //120
                //-------------------------------------------------------------------

                // 2014 그 밖의 소득공제 계 
                vGDColumnIndex = pGDColumn[225];
                vXLColumnIndex = pXLColumn[225];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;    //121
                //-------------------------------------------------------------------

                // (49)소득공제 종합한도 초과액
                vGDColumnIndex = pGDColumn[148];
                vXLColumnIndex = pXLColumn[148];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액
                vGDColumnIndex = pGDColumn[162];
                vXLColumnIndex = pXLColumn[162];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

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

        #region ----- XLLINE13 -----

        private int XLLine13(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, object pPrintDate, string pPrint_Type, object pPrint_Type_Desc)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            //System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //----[ 1 page ]------------------------------------------------------------------------------------------------------

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 거주 구분(거주자1/거주자2)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //거주자1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //거주자 2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //거주자1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //거주자 2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 거주지국
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 거주지국코드
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 내외국인 구분(내국인1/외국인9)
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //내국인1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //외국인9이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //내국인1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //외국인9이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 외국인단일세율적용
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "Y") //여1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //부2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 2), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //여1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //부2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 3), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 출력 용도 구분
                vXLColumnIndex = 14;
                vObject = pPrint_Type_Desc;
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

                // 국적
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 국적코드
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 세대주 구분(세대주1/세대원2)
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //세대주1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //세대원2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //세대주1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //세대원2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 연말정산구분
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "계속근로") //계속근로1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //중도퇴사2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //계속근로1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //중도퇴사2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 법인명(상호)
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 대표자(성명)    
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업자등록번호
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 소재지(주소)
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 성명
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                string sName = vConvertString;

                // 주민번호
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                string sPersonNumber = vConvertString;

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주소
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 종전
                vGDColumnIndex = pGDColumn[164];
                vXLColumnIndex = pXLColumn[164];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)근무처명
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1근무처명
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2근무처명
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3근무처명
                vGDColumnIndex = pGDColumn[165];
                vXLColumnIndex = pXLColumn[165];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)사업자번호
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1사업잡번호
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
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

                // 종(전)2사업잡번호
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3사업잡번호
                vGDColumnIndex = pGDColumn[166];
                vXLColumnIndex = pXLColumn[166];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)근무기간
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1근무기간
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2근무기간
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3근무기간
                vGDColumnIndex = pGDColumn[167];
                vXLColumnIndex = pXLColumn[167];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)감면기간
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1감면기간
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2감면기간
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3감면기간
                vGDColumnIndex = pGDColumn[168];
                vXLColumnIndex = pXLColumn[168];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)급여 
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1급여
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2급여
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3급여
                vGDColumnIndex = pGDColumn[169];
                vXLColumnIndex = pXLColumn[169];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)상여
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1상여
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2상여
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3상여
                vGDColumnIndex = pGDColumn[170];
                vXLColumnIndex = pXLColumn[170];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)인정상여 
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1인정상여
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2인정상여
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3인정상여
                vGDColumnIndex = pGDColumn[171];
                vXLColumnIndex = pXLColumn[171];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)주식매수선택권
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1주식매수선택권
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2주식매수선택권
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3주식매수선택권
                vGDColumnIndex = pGDColumn[172];
                vXLColumnIndex = pXLColumn[172];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)우리사주조합인출금
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1우리사주조합인출금
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2우리사주조합인출금
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3우리사주조합인출금
                vGDColumnIndex = pGDColumn[173];
                vXLColumnIndex = pXLColumn[173];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)임원퇴직소득금액 한도초과액
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1임원퇴직소득금액 한도초과액
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2임원퇴직소득금액 한도초과액
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3임원퇴직소득금액 한도초과액
                vGDColumnIndex = pGDColumn[174];
                vXLColumnIndex = pXLColumn[174];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주(현)계
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1계
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2계
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)3계
                vGDColumnIndex = pGDColumn[175];
                vXLColumnIndex = pXLColumn[175];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //--------------------------------------------------------------------------------------------------------------------
                // 비과세 및 감면 소득 명세
                //--------------------------------------------------------------------------------------------------------------------

                // 비과세_주(현)국외근로
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1국외근로
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2국외근로
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3국외근로
                vGDColumnIndex = pGDColumn[176];
                vXLColumnIndex = pXLColumn[176];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)야간근로수당
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1야간근로수당
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2야간근로수당
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3야간근로수당
                vGDColumnIndex = pGDColumn[177];
                vXLColumnIndex = pXLColumn[177];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)출산/보육수당
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1출산/보육수당
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2출산/보육수당
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2출산/보육수당
                vGDColumnIndex = pGDColumn[178];
                vXLColumnIndex = pXLColumn[178];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)연구보조비
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1연구보조비
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2연구보조비
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2연구보조비
                vGDColumnIndex = pGDColumn[179];
                vXLColumnIndex = pXLColumn[179];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // 중소기업에 취업하는 청년에 대한 소득세 감면 존재
                vGDColumnIndex = pGDColumn[187];
                vXLColumnIndex = pXLColumn[187];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }

                // 중소기업에 취업하는 청년에 대한 소득세 감면1
                vGDColumnIndex = pGDColumn[188];
                vXLColumnIndex = pXLColumn[188];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }


                // 중소기업에 취업하는 청년에 대한 소득세 감면2
                vGDColumnIndex = pGDColumn[189];
                vXLColumnIndex = pXLColumn[189];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }

                // 중소기업에 취업하는 청년에 대한 소득세 감면3
                vGDColumnIndex = pGDColumn[190];
                vXLColumnIndex = pXLColumn[190];

                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                if (vObject != null)
                {
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
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 비과세_주(현)수련보조수당
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1수련보조수당
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2수련보조수당
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3수련보조수당
                vGDColumnIndex = pGDColumn[180];
                vXLColumnIndex = pXLColumn[180];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)비과세소득 계
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1비과세소득 계
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2비과세소득 계
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3비과세소득 계
                vGDColumnIndex = pGDColumn[181];
                vXLColumnIndex = pXLColumn[181];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)감면소득계
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1감면소득계
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2감면소득계
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)3감면소득계
                vGDColumnIndex = pGDColumn[182];
                vXLColumnIndex = pXLColumn[182];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                //--------------------------------------------------------------------------------------------------------------------
                // 세액 명세
                //--------------------------------------------------------------------------------------------------------------------

                // 결정세액_소득세
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_지방소득세   
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_농특세
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)1사업자번호 
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1소득세
                vGDColumnIndex = pGDColumn[74];
                vXLColumnIndex = pXLColumn[74];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1지방소득세
                vGDColumnIndex = pGDColumn[75];
                vXLColumnIndex = pXLColumn[75];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1농특세
                vGDColumnIndex = pGDColumn[76];
                vXLColumnIndex = pXLColumn[76];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)2사업자번호 
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2소득세
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2지방소득세
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2농특세
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)3사업자번호 
                vGDColumnIndex = pGDColumn[183];
                vXLColumnIndex = pXLColumn[183];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)3소득세
                vGDColumnIndex = pGDColumn[184];
                vXLColumnIndex = pXLColumn[184];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)3지방소득세
                vGDColumnIndex = pGDColumn[185];
                vXLColumnIndex = pXLColumn[185];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)3농특세
                vGDColumnIndex = pGDColumn[186];
                vXLColumnIndex = pXLColumn[186];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_주(현)소득세 
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_주(현)지방소득세
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //기납부세액_주(현)농특세
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 차감징수세액_소득세 
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_지방소득세
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_농특세
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[86];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 5;
                //-------------------------------------------------------------------

                // 날짜
                vXLColumnIndex = 28;
                vObject = pPrintDate;
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

                // 징수의무자
                vXLColumnIndex = 23;
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WITHHOLDING_OWNER");
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 9;
                //-------------------------------------------------------------------

                //----[ 2 page ]------------------------------------------------------------------------------------------------------

                // 2page 상단에 소득자 성명 및 주민번호 출력 표시되는 부분
                string sPrintPersinInfo = sName + "(" + sPersonNumber + ")";
                mPrinting.XLSetCell(vXLine, 24, sPrintPersinInfo);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 총급여
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[87];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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


                // 개인연금저축소득공제
                vGDColumnIndex = pGDColumn[137];
                vXLColumnIndex = pXLColumn[137];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 근로소득공제
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[88];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 소기업/소상공인 공제부금 소득공제
                vGDColumnIndex = pGDColumn[138];
                vXLColumnIndex = pXLColumn[138];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 근로소득금액
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[89];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(본인)
                vGDColumnIndex = pGDColumn[90];
                vXLColumnIndex = pXLColumn[90];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 청약저축
                vGDColumnIndex = pGDColumn[139];
                vXLColumnIndex = pXLColumn[139];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(배우자)
                vGDColumnIndex = pGDColumn[91];
                vXLColumnIndex = pXLColumn[91];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주택청약종합저축
                vGDColumnIndex = pGDColumn[140];
                vXLColumnIndex = pXLColumn[140];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(부양인원 - 인원)  
                vGDColumnIndex = pGDColumn[92];
                vXLColumnIndex = pXLColumn[92];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기본(부양인원 - 금액) 
                vGDColumnIndex = pGDColumn[93];
                vXLColumnIndex = pXLColumn[93];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 근로자주택마련저축
                vGDColumnIndex = pGDColumn[141];
                vXLColumnIndex = pXLColumn[141];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 추가공제(경로수 - 인원)
                vGDColumnIndex = pGDColumn[94];
                vXLColumnIndex = pXLColumn[94];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(경로수 - 금액)
                vGDColumnIndex = pGDColumn[95];
                vXLColumnIndex = pXLColumn[95];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(장애인 - 인원)
                vGDColumnIndex = pGDColumn[96];
                vXLColumnIndex = pXLColumn[96];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(장애인 - 금액)
                vGDColumnIndex = pGDColumn[97];
                vXLColumnIndex = pXLColumn[97];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 투자조합출자등소득공제
                vGDColumnIndex = pGDColumn[142];
                vXLColumnIndex = pXLColumn[142];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(부녀세대)
                vGDColumnIndex = pGDColumn[98];
                vXLColumnIndex = pXLColumn[98];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 신용카드등소득공제
                vGDColumnIndex = pGDColumn[143];
                vXLColumnIndex = pXLColumn[143];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(자녀양육 - 인원)    
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(자녀양육 - 금액)
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 우리사주조합소득공제
                vGDColumnIndex = pGDColumn[144];
                vXLColumnIndex = pXLColumn[144];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(출산입양 - 인원) 
                vGDColumnIndex = pGDColumn[101];
                vXLColumnIndex = pXLColumn[101];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(출산입양 - 금액)
                vGDColumnIndex = pGDColumn[102];
                vXLColumnIndex = pXLColumn[102];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 추가공제(한부모가족 - 금액)
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //  고용유지중소기업근로자
                vGDColumnIndex = pGDColumn[145];
                vXLColumnIndex = pXLColumn[145];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 다자녀공제(인원)  
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 다자녀공제(금액)  
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 목돈 안드는 전세이자상환액 
                vGDColumnIndex = pGDColumn[146];
                vXLColumnIndex = pXLColumn[146];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 국민연금보험료공제
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 공무원연금
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 그 밖의 소득공제 계
                vGDColumnIndex = pGDColumn[147];
                vXLColumnIndex = pXLColumn[147];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 군인연금
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 특별공제 종합한도 초과액
                vGDColumnIndex = pGDColumn[148];
                vXLColumnIndex = pXLColumn[148];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 사립합교교직원연금
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 특종합소득 과세표준
                vGDColumnIndex = pGDColumn[149];
                vXLColumnIndex = pXLColumn[149];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 별정우체국연금
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 산출세액
                vGDColumnIndex = pGDColumn[150];
                vXLColumnIndex = pXLColumn[150];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 과학기술인공제
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 소득세법
                vGDColumnIndex = pGDColumn[151];
                vXLColumnIndex = pXLColumn[151];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 근로자퇴직급여 보장법에 따른 퇴직연금
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 「조세특례제한법」(<53>-1제외)
                vGDColumnIndex = pGDColumn[152];
                vXLColumnIndex = pXLColumn[152];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 연금저축
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 「조세특례제한법」 제30조
                vGDColumnIndex = pGDColumn[153];
                vXLColumnIndex = pXLColumn[153];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 건강보험료(노인장기요양보험료 포함)
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 조세조약
                vGDColumnIndex = pGDColumn[154];
                vXLColumnIndex = pXLColumn[154];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 고용보험료
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 보장성보험
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 장애인전용
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 세 액 감 면 계
                vGDColumnIndex = pGDColumn[155];
                vXLColumnIndex = pXLColumn[155];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 의료비-장애인
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 의료비-기타
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 교육비-장애인
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 근로소득
                vGDColumnIndex = pGDColumn[156];
                vXLColumnIndex = pXLColumn[156];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 교육비-기타
                vGDColumnIndex = pGDColumn[121];
                vXLColumnIndex = pXLColumn[121];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 주택임차차입금원리금상환액-대출기관
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 납세조합공제
                vGDColumnIndex = pGDColumn[157];
                vXLColumnIndex = pXLColumn[157];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 주택임차차입금원리금상환액-거주자
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // 주택차입금
                vGDColumnIndex = pGDColumn[158];
                vXLColumnIndex = pXLColumn[158];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 월세액
                vGDColumnIndex = pGDColumn[124];
                vXLColumnIndex = pXLColumn[124];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 기부정치자금
                vGDColumnIndex = pGDColumn[159];
                vXLColumnIndex = pXLColumn[159];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 장기주택저당차입금이자상환액 - 2011이전 (15년미만)
                vGDColumnIndex = pGDColumn[125];
                vXLColumnIndex = pXLColumn[125];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 외국납부
                vGDColumnIndex = pGDColumn[160];
                vXLColumnIndex = pXLColumn[160];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 장기주택저당차입금이자상환액 - 2011이전 (15년~29년)
                vGDColumnIndex = pGDColumn[126];
                vXLColumnIndex = pXLColumn[126];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //장기주택저당차입금이자상환액 - 2011이전 (30년 이상)
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 장기주택저당차입금이자상환액 - 2012이후 (고정금리비거치 상환대출)
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 장기주택저당차입금이자상환액 - 2012이후 (기타)
                vGDColumnIndex = pGDColumn[129];
                vXLColumnIndex = pXLColumn[129];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 정치자금기부금
                vGDColumnIndex = pGDColumn[130];
                vXLColumnIndex = pXLColumn[130];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 법정기부금
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 우리사주조합기부금      
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //지정기부금     
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 계
                vGDColumnIndex = pGDColumn[134];
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 세액공제 계
                vGDColumnIndex = pGDColumn[161];
                vXLColumnIndex = pXLColumn[161];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 표준공제
                vGDColumnIndex = pGDColumn[135];
                vXLColumnIndex = pXLColumn[135];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 차감소득금액
                vGDColumnIndex = pGDColumn[136];
                vXLColumnIndex = pXLColumn[136];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액
                vGDColumnIndex = pGDColumn[162];
                vXLColumnIndex = pXLColumn[162];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

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

        #region ----- XLLINE12 -----

        private int XLLine12(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, object vPrintDate, object vPrintType)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            //System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //----[ 1 page ]------------------------------------------------------------------------------------------------------

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 거주 구분(거주자1/거주자2)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //거주자1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //거주자 2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //거주자1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //거주자 2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 내외국인 구분(내국인1/외국인9)
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //내국인1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //외국인9이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //내국인1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //외국인9이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 외국인단일세율적용
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "Y") //여1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //부2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 3), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //여1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //부2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 3), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 출력 용도 구분
                vXLColumnIndex = 14;
                vObject = vPrintType;
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 세대주 구분(세대주1/세대원2)
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //세대주1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //세대원2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //세대주1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //세대원2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 연말정산구분
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "계속근로") //계속근로1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //중도퇴사2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //계속근로1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //중도퇴사2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 법인명(상호)
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 대표자(성명)    
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업자등록번호
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 소재지(주소)
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 성명
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                string sName = vConvertString;

                // 주민번호
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                string sPersonNumber = vConvertString;

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주소
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주(현)근무처명
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1근무처명
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2근무처명
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)사업자번호
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1사업잡번호
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
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

                // 종(전)2사업잡번호
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)근무기간
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1근무기간
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2근무기간
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)감면기간
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1감면기간
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2감면기간
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)급여 
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1급여
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2급여
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)상여
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1상여
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2상여
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)인정상여 
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1인정상여
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2인정상여
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)주식매수선택권
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1주식매수선택권
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2주식매수선택권
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)우리사주조합인출금
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1우리사주조합인출금
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2우리사주조합인출금
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // 주(현)계
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1계
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2계
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //--------------------------------------------------------------------------------------------------------------------
                // II 비과세 및 감면 소득 명세
                //--------------------------------------------------------------------------------------------------------------------

                // 비과세_주(현)국외근로
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1국외근로
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2국외근로
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)야간근로수당
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1야간근로수당
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2야간근로수당
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)출산/보육수당
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1출산/보육수당
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2출산/보육수당
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)외국인근로자
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1외국인근로자
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2외국인근로자
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 6;
                //-------------------------------------------------------------------

                // 비과세_주(현)비과세소득계
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1비과세소득계
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2비과세소득계
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)감면소득계
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1감면소득계
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2감면소득계
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                //--------------------------------------------------------------------------------------------------------------------
                // III 세액 명세
                //--------------------------------------------------------------------------------------------------------------------

                // 결정세액_소득세
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_지방소득세   
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_농특세
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_계
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)1사업자번호 
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1소득세
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1지방소득세
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1농특세
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1계
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)2사업자번호 
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2소득세
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2지방소득세
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2농특세
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2계
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_주(현)소득세 
                vGDColumnIndex = pGDColumn[74];
                vXLColumnIndex = pXLColumn[74];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_주(현)지방소득세
                vGDColumnIndex = pGDColumn[75];
                vXLColumnIndex = pXLColumn[75];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //기납부세액_주(현)농특세
                vGDColumnIndex = pGDColumn[76];
                vXLColumnIndex = pXLColumn[76];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_주(현)계
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 차감징수세액_소득세 
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_지방소득세
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_농특세
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_계
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 6;
                //-------------------------------------------------------------------

                // 날짜
                vXLColumnIndex = 28;
                vObject = vPrintDate;
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

                // 징수의무자
                vXLColumnIndex = 23;
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WITHHOLDING_OWNER");
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 7;
                //-------------------------------------------------------------------

                //----[ 2 page ]------------------------------------------------------------------------------------------------------

                // 2page 상단에 소득자 성명 및 주민번호 출력 표시되는 부분
                string sPrintPersinInfo = sName + "(" + sPersonNumber + ")";
                mPrinting.XLSetCell(vXLine, 24, sPrintPersinInfo);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 총급여
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 개인연금저축소득공제
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 근로소득공제
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 연금저축소득공제
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 근로소득금액
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[86];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 소기업/소상공인 공제부금 소득공제
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[87];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(본인)
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[88];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 청약저축
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[89];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(배우자)
                vGDColumnIndex = pGDColumn[90];
                vXLColumnIndex = pXLColumn[90];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주택청약종합저축
                vGDColumnIndex = pGDColumn[91];
                vXLColumnIndex = pXLColumn[91];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(부양인원 - 인원)  
                vGDColumnIndex = pGDColumn[92];
                vXLColumnIndex = pXLColumn[92];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기본(부양인원 - 금액) 
                vGDColumnIndex = pGDColumn[93];
                vXLColumnIndex = pXLColumn[93];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 장기주택마련저축
                vGDColumnIndex = pGDColumn[94];
                vXLColumnIndex = pXLColumn[94];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(경로수 - 인원)
                vGDColumnIndex = pGDColumn[95];
                vXLColumnIndex = pXLColumn[95];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(경로수 - 금액)
                vGDColumnIndex = pGDColumn[96];
                vXLColumnIndex = pXLColumn[96];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 근로자주택마련저축
                vGDColumnIndex = pGDColumn[97];
                vXLColumnIndex = pXLColumn[97];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(장애인 - 인원)
                vGDColumnIndex = pGDColumn[98];
                vXLColumnIndex = pXLColumn[98];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(장애인 - 금액)
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 투자조합출자등 소득공제
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(부녀세대)
                vGDColumnIndex = pGDColumn[101];
                vXLColumnIndex = pXLColumn[101];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 신용카드등 소득공제
                vGDColumnIndex = pGDColumn[102];
                vXLColumnIndex = pXLColumn[102];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(자녀양육 - 인원)    
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(자녀양육 - 금액)
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 우리사주출자
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(출산입양 - 인원) 
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(출산입양 - 금액)
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 장기주식형저축
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 고용유지중소기업소득공제
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 다자녀공제(인원)  
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 다자녀공제(금액)  
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 국민연금보험료공제
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------
                // 그 밖의 소득공제 계
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 종합소득 과세표준
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 산출세액
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 소득세법
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 근로자퇴직연금소득공제
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 조세특례제한법
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 건강보험료
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 고용보험료
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 보장성 보험
                vGDColumnIndex = pGDColumn[121];
                vXLColumnIndex = pXLColumn[121];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 장애인 전용
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 의료비
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 세액감면 계
                vGDColumnIndex = pGDColumn[124];
                vXLColumnIndex = pXLColumn[124];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 교육비
                vGDColumnIndex = pGDColumn[125];
                vXLColumnIndex = pXLColumn[125];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 근로소득
                vGDColumnIndex = pGDColumn[126];
                vXLColumnIndex = pXLColumn[126];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 주택임차차입금 원리금 상환액
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 납세조합 공제
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;  //2013
                //-------------------------------------------------------------------
                // 월세액
                vGDColumnIndex = pGDColumn[129];
                vXLColumnIndex = pXLColumn[129];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주택차입금
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 장기주택저당 차입금 이자 상환액(2011이전)
                vGDColumnIndex = pGDColumn[130];
                vXLColumnIndex = pXLColumn[130];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기부 정치자금
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 장기주택저당 차입금 이자 상환액(2012이후)
                vGDColumnIndex = pGDColumn[140];
                vXLColumnIndex = pXLColumn[140];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 기부금
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 외국납부
                vGDColumnIndex = pGDColumn[134];
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 계
                vGDColumnIndex = pGDColumn[135];
                vXLColumnIndex = pXLColumn[135];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 표준공제
                vGDColumnIndex = pGDColumn[136];
                vXLColumnIndex = pXLColumn[136];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 세액공제 계
                vGDColumnIndex = pGDColumn[137];
                vXLColumnIndex = pXLColumn[137];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 차감소득금액
                vGDColumnIndex = pGDColumn[138];
                vXLColumnIndex = pXLColumn[138];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액
                vGDColumnIndex = pGDColumn[139];
                vXLColumnIndex = pXLColumn[139];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
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

        #region ----- XLLINE11 -----

        private int XLLine11(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, object vPrintDate, object vPrintType)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            //System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //----[ 1 page ]------------------------------------------------------------------------------------------------------

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 거주 구분(거주자1/거주자2)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //거주자1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //거주자 2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //거주자1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //거주자 2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 내외국인 구분(내국인1/외국인9)
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //내국인1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //외국인9이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //내국인1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //외국인9이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 외국인단일세율적용
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "Y") //여1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //부2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 3), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //여1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //부2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 3), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 출력 용도 구분
                vXLColumnIndex = 14;
                vObject = vPrintType;
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 세대주 구분(세대주1/세대원2)
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "1") //세대주1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //세대원2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //세대주1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //세대원2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 연말정산구분
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "계속근로") //계속근로1이면,
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else //중도퇴사2이면,
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    //계속근로1이면,
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //중도퇴사2이면,
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 5), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 법인명(상호)
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 대표자(성명)    
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업자등록번호
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 소재지(주소)
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 성명
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                string sName = vConvertString;

                // 주민번호
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                string sPersonNumber = vConvertString;

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주소
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주(현)근무처명
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1근무처명
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2근무처명
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)사업자번호
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1사업잡번호
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
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

                // 종(전)2사업잡번호
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)근무기간
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1근무기간
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2근무기간
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)감면기간
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1감면기간
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2감면기간
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)급여 
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1급여
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2급여
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)상여
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1상여
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2상여
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)인정상여 
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1인정상여
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2인정상여
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)주식매수선택권
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1주식매수선택권
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2주식매수선택권
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주(현)우리사주조합인출금
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1우리사주조합인출금
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2우리사주조합인출금
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // 주(현)계
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)1계
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 종(전)2계
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //--------------------------------------------------------------------------------------------------------------------
                // II 비과세 및 감면 소득 명세
                //--------------------------------------------------------------------------------------------------------------------

                // 비과세_주(현)국외근로
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1국외근로
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2국외근로
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)야간근로수당
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1야간근로수당
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2야간근로수당
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)출산/보육수당
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1출산/보육수당
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2출산/보육수당
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)외국인근로자
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1외국인근로자
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2외국인근로자
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 6;
                //-------------------------------------------------------------------

                // 비과세_주(현)비과세소득계
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1비과세소득계
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2비과세소득계
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 비과세_주(현)감면소득계
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)1감면소득계
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 비과세_종(전)2감면소득계
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                //--------------------------------------------------------------------------------------------------------------------
                // III 세액 명세
                //--------------------------------------------------------------------------------------------------------------------

                // 결정세액_소득세
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_지방소득세   
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_농특세
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액_계
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)1사업자번호 
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1소득세
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1지방소득세
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1농특세
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)1계
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 기납부세액_종(전)2사업자번호 
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2소득세
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2지방소득세
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2농특세
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_종(전)2계
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_주(현)소득세 
                vGDColumnIndex = pGDColumn[74];
                vXLColumnIndex = pXLColumn[74];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_주(현)지방소득세
                vGDColumnIndex = pGDColumn[75];
                vXLColumnIndex = pXLColumn[75];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                //기납부세액_주(현)농특세
                vGDColumnIndex = pGDColumn[76];
                vXLColumnIndex = pXLColumn[76];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기납부세액_주(현)계
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 차감징수세액_소득세 
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_지방소득세
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_농특세
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 차감징수세액_계
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 6;
                //-------------------------------------------------------------------

                // 날짜
                vXLColumnIndex = 28;
                vObject = vPrintDate;
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

                // 징수의무자
                vXLColumnIndex = 23;
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WITHHOLDING_OWNER");
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 7;
                //-------------------------------------------------------------------

                //----[ 2 page ]------------------------------------------------------------------------------------------------------

                // 2page 상단에 소득자 성명 및 주민번호 출력 표시되는 부분
                string sPrintPersinInfo = sName + "(" + sPersonNumber + ")";
                mPrinting.XLSetCell(vXLine, 24, sPrintPersinInfo);

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 총급여
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 개인연금저축소득공제
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 근로소득공제
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 연금저축소득공제
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 근로소득금액
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[86];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 소기업/소상공인 공제부금 소득공제
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[87];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(본인)
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[88];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 청약저축
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[89];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(배우자)
                vGDColumnIndex = pGDColumn[90];
                vXLColumnIndex = pXLColumn[90];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주택청약종합저축
                vGDColumnIndex = pGDColumn[91];
                vXLColumnIndex = pXLColumn[91];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 기본(부양인원 - 인원)  
                vGDColumnIndex = pGDColumn[92];
                vXLColumnIndex = pXLColumn[92];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기본(부양인원 - 금액) 
                vGDColumnIndex = pGDColumn[93];
                vXLColumnIndex = pXLColumn[93];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 장기주택마련저축
                vGDColumnIndex = pGDColumn[94];
                vXLColumnIndex = pXLColumn[94];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(경로수 - 인원)
                vGDColumnIndex = pGDColumn[95];
                vXLColumnIndex = pXLColumn[95];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(경로수 - 금액)
                vGDColumnIndex = pGDColumn[96];
                vXLColumnIndex = pXLColumn[96];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 근로자주택마련저축
                vGDColumnIndex = pGDColumn[97];
                vXLColumnIndex = pXLColumn[97];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(장애인 - 인원)
                vGDColumnIndex = pGDColumn[98];
                vXLColumnIndex = pXLColumn[98];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(장애인 - 금액)
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 투자조합출자등 소득공제
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(부녀세대)
                vGDColumnIndex = pGDColumn[101];
                vXLColumnIndex = pXLColumn[101];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 신용카드등 소득공제
                vGDColumnIndex = pGDColumn[102];
                vXLColumnIndex = pXLColumn[102];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(자녀양육 - 인원)    
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(자녀양육 - 금액)
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 우리사주출자
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 추가공제(출산입양 - 인원) 
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 추가공제(출산입양 - 금액)
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 장기주식형저축
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 고용유지중소기업소득공제
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 다자녀공제(인원)  
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 다자녀공제(금액)  
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 국민연금보험료공제
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------
                // 그 밖의 소득공제 계
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 종합소득 과세표준
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 산출세액
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 소득세법
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 근로자퇴직연금소득공제
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 조세특례제한법
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 건강보험료
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 고용보험료
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 보장성 보험
                vGDColumnIndex = pGDColumn[121];
                vXLColumnIndex = pXLColumn[121];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 장애인 전용
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 의료비
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 세액감면 계
                vGDColumnIndex = pGDColumn[124];
                vXLColumnIndex = pXLColumn[124];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 교육비
                vGDColumnIndex = pGDColumn[125];
                vXLColumnIndex = pXLColumn[125];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 근로소득
                vGDColumnIndex = pGDColumn[126];
                vXLColumnIndex = pXLColumn[126];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 주택임차차입금 원리금 상환액
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 납세조합 공제
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 월세액
                vGDColumnIndex = pGDColumn[129];
                vXLColumnIndex = pXLColumn[129];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 장기주택저당 차입금 이자 상환액
                vGDColumnIndex = pGDColumn[130];
                vXLColumnIndex = pXLColumn[130];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주택차입금
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------
                // 기부금
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 기부 정치자금
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 외국납부
                vGDColumnIndex = pGDColumn[134];
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                // 계
                vGDColumnIndex = pGDColumn[135];
                vXLColumnIndex = pXLColumn[135];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 표준공제
                vGDColumnIndex = pGDColumn[136];
                vXLColumnIndex = pXLColumn[136];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 세액공제 계
                vGDColumnIndex = pGDColumn[137];
                vXLColumnIndex = pXLColumn[137];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 차감소득금액
                vGDColumnIndex = pGDColumn[138];
                vXLColumnIndex = pXLColumn[138];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 결정세액
                vGDColumnIndex = pGDColumn[139];
                vXLColumnIndex = pXLColumn[139];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 3;
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

        #region ----- XLLINE14_2 : 부양가족내역 -----

        private int XLLINE14_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_SUPPORT_FAMILY, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");
                if (pGridRow == 0)
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 14;
                    //-------------------------------------------------------------------
                    // 다자녀 인원 수
                    vGDColumnIndex = pGDColumn[0];
                    vXLColumnIndex = pXLColumn[0];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                }

                //----[ 3 page ]------------------------------------------------------------------------------------------------------
                if (pGridRow == -1)
                {
                    //-------------------------------------------------------------------
                    vXLine = 134;
                    //-------------------------------------------------------------------

                    //// 기본공제
                    //vXLColumnIndex = pXLColumn[31];
                    //if (iString.ISDecimaltoZero(mBASE_COUNT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mBASE_COUNT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, mBASE_COUNT);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 경로우대
                    //vXLColumnIndex = pXLColumn[32];
                    //if (iString.ISDecimaltoZero(mOLD_COUNT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mOLD_COUNT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 출산/입양양육
                    //vXLColumnIndex = pXLColumn[33];
                    //if (iString.ISDecimaltoZero(mBIRTH_COUNT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mBIRTH_COUNT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-보험료
                    //vXLColumnIndex = pXLColumn[37];
                    //if (iString.ISDecimaltoZero(mINSURE_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mINSURE_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-의료비
                    //vXLColumnIndex = pXLColumn[38];
                    //if (iString.ISDecimaltoZero(mMEDICAL_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mMEDICAL_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-교육비
                    //vXLColumnIndex = pXLColumn[39];
                    //if (iString.ISDecimaltoZero(mEDU_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mEDU_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-신용카드
                    //vXLColumnIndex = pXLColumn[40];
                    //if (iString.ISDecimaltoZero(mCREDIT_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mCREDIT_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-직불카드
                    //vXLColumnIndex = pXLColumn[41];
                    //if (iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mCHECK_CREDIT_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-학원비지로납부액
                    //vXLColumnIndex = pXLColumn[42];
                    //if (iString.ISDecimaltoZero(mCASH_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mCASH_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-현금영수증
                    //vXLColumnIndex = pXLColumn[43];
                    //if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mDONAT_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-전통시장사용액 
                    //vXLColumnIndex = pXLColumn[44];
                    //if (iString.ISDecimaltoZero(mTRAD_MARKET_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mTRAD_MARKET_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 국세청-기부금
                    //vXLColumnIndex = pXLColumn[45];
                    //if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mDONAT_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}


                    ////-------------------------------------------------------------------
                    //vXLine = vXLine + 1;
                    ////-------------------------------------------------------------------

                    //// 부녀자
                    //vXLColumnIndex = pXLColumn[34];
                    //if (iString.ISDecimaltoZero(mBASE_COUNT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mBASE_COUNT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, mBASE_COUNT);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 장애인
                    //vXLColumnIndex = pXLColumn[35];
                    //if (iString.ISDecimaltoZero(mOLD_COUNT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mOLD_COUNT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 6세이하
                    //vXLColumnIndex = pXLColumn[36];
                    //if (iString.ISDecimaltoZero(mBIRTH_COUNT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mBIRTH_COUNT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 기타-보장성보험료
                    //vXLColumnIndex = pXLColumn[48];
                    //if (iString.ISDecimaltoZero(mINSURE_ETC_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mINSURE_ETC_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 기타-의료비
                    //vXLColumnIndex = pXLColumn[49];
                    //if (iString.ISDecimaltoZero(mMEDICAL_ETC_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mMEDICAL_ETC_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 기타-교육비
                    //vXLColumnIndex = pXLColumn[50];
                    //if (iString.ISDecimaltoZero(mEDU_ETC_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mEDU_ETC_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 기타-신용카드
                    //vXLColumnIndex = pXLColumn[51];
                    //if (iString.ISDecimaltoZero(mCREDIT_ETC_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mCREDIT_ETC_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 기타-직불카드
                    //vXLColumnIndex = pXLColumn[52];
                    //if (iString.ISDecimaltoZero(mCHECK_CREDIT_ETC_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mCHECK_CREDIT_ETC_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 기타-학원비지로납부액
                    //vXLColumnIndex = pXLColumn[53];
                    //if (iString.ISDecimaltoZero(mACADE_GIRO_ETC_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mACADE_GIRO_ETC_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 기타-현금영수증
                    ////vXLColumnIndex = pXLColumn[54];
                    ////if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    ////{
                    ////    vConvertString = string.Format("{0}", mDONAT_AMT);
                    ////    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    ////}
                    ////else
                    ////{
                    ////    vConvertString = string.Empty;
                    ////    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    ////}

                    //// 기타-전통시장사용액 
                    //vXLColumnIndex = pXLColumn[55];
                    //if (iString.ISDecimaltoZero(mTRAD_MARKET_ETC_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mTRAD_MARKET_ETC_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    //// 기타-기부금
                    //vXLColumnIndex = pXLColumn[56];
                    //if (iString.ISDecimaltoZero(mDONAT_ETC_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mDONAT_ETC_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}



                    vXLine = pXLine;
                }
                else
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 관계코드
                    vGDColumnIndex = pGDColumn[1];
                    vXLColumnIndex = pXLColumn[1];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 성명
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기본공제
                    vGDColumnIndex = pGDColumn[3];
                    vXLColumnIndex = pXLColumn[3];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 경로우대
                    vGDColumnIndex = pGDColumn[4];
                    vXLColumnIndex = pXLColumn[4];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 출산/입양양육
                    vGDColumnIndex = pGDColumn[5];
                    vXLColumnIndex = pXLColumn[5];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 국세청-건강/고용보험료
                    vGDColumnIndex = pGDColumn[6];
                    vXLColumnIndex = pXLColumn[6];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mINSURE_AMT = iString.ISDecimaltoZero(mINSURE_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-보험료 
                    vGDColumnIndex = pGDColumn[7];
                    vXLColumnIndex = pXLColumn[7];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mMEDICAL_AMT = iString.ISDecimaltoZero(mMEDICAL_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-장애인보험료 
                    vGDColumnIndex = pGDColumn[8];
                    vXLColumnIndex = pXLColumn[8];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mEDU_AMT = iString.ISDecimaltoZero(mEDU_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-의료비
                    vGDColumnIndex = pGDColumn[9];
                    vXLColumnIndex = pXLColumn[9];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mCREDIT_AMT = iString.ISDecimaltoZero(mCREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-교육비
                    vGDColumnIndex = pGDColumn[10];
                    vXLColumnIndex = pXLColumn[10];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-신용카드
                    vGDColumnIndex = pGDColumn[11];
                    vXLColumnIndex = pXLColumn[11];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-직불카드
                    vGDColumnIndex = pGDColumn[12];
                    vXLColumnIndex = pXLColumn[12];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-현금
                    vGDColumnIndex = pGDColumn[13];
                    vXLColumnIndex = pXLColumn[13];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-전통시장
                    vGDColumnIndex = pGDColumn[14];
                    vXLColumnIndex = pXLColumn[14];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 국세청-대중교통
                    vGDColumnIndex = pGDColumn[15];
                    vXLColumnIndex = pXLColumn[15];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 국세청-기부금
                    vGDColumnIndex = pGDColumn[16];
                    vXLColumnIndex = pXLColumn[16];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        if (vConvertDecimal != 0)
                        {
                            vConvertString = string.Format("{0}", vConvertDecimal);
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                        }
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //mDONAT_AMT = iString.ISDecimaltoZero(mDONAT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 국가타입
                    vGDColumnIndex = pGDColumn[17];
                    vXLColumnIndex = pXLColumn[17];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 주민번호
                    vGDColumnIndex = pGDColumn[18];
                    vXLColumnIndex = pXLColumn[18];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 부녀자
                    vGDColumnIndex = pGDColumn[19];
                    vXLColumnIndex = pXLColumn[19];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mWOMAN_COUNT = iString.ISDecimaltoZero(mWOMAN_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[30]), 0);

                    // 한부모
                    vGDColumnIndex = pGDColumn[20];
                    vXLColumnIndex = pXLColumn[20];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mWOMAN_COUNT = iString.ISDecimaltoZero(mWOMAN_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[30]), 0);

                    // 장애인
                    vGDColumnIndex = pGDColumn[21];
                    vXLColumnIndex = pXLColumn[21];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mDISABILITY_COUNT = iString.ISDecimaltoZero(mDISABILITY_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[28]), 0);

                    // 자녀양육(6세이하)
                    vGDColumnIndex = pGDColumn[22];
                    vXLColumnIndex = pXLColumn[22];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHILD_COUNT = iString.ISDecimaltoZero(mCHILD_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[29]), 0);

                    // 기타-건강고용보험료
                    vGDColumnIndex = pGDColumn[23];
                    vXLColumnIndex = pXLColumn[23];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기타-보험료
                    vGDColumnIndex = pGDColumn[24];
                    vXLColumnIndex = pXLColumn[24];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mINSURE_ETC_AMT = iString.ISDecimaltoZero(mINSURE_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-장애인보험료
                    vGDColumnIndex = pGDColumn[25];
                    vXLColumnIndex = pXLColumn[25];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mMEDICAL_ETC_AMT = iString.ISDecimaltoZero(mMEDICAL_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-의료비
                    vGDColumnIndex = pGDColumn[26];
                    vXLColumnIndex = pXLColumn[26];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mEDU_ETC_AMT = iString.ISDecimaltoZero(mEDU_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-교육비
                    vGDColumnIndex = pGDColumn[27];
                    vXLColumnIndex = pXLColumn[27];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCREDIT_ETC_AMT = iString.ISDecimaltoZero(mCREDIT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-신용카드
                    vGDColumnIndex = pGDColumn[28];
                    vXLColumnIndex = pXLColumn[28];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_ETC_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-직불카드
                    vGDColumnIndex = pGDColumn[29];
                    vXLColumnIndex = pXLColumn[29];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mACADE_GIRO_ETC_AMT = iString.ISDecimaltoZero(mACADE_GIRO_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-현금
                    vGDColumnIndex = pGDColumn[30];
                    vXLColumnIndex = pXLColumn[30];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기타-전통시장
                    vGDColumnIndex = pGDColumn[31];
                    vXLColumnIndex = pXLColumn[31];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mTRAD_MARKET_ETC_AMT = iString.ISDecimaltoZero(mTRAD_MARKET_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-대중교통
                    vGDColumnIndex = pGDColumn[32];
                    vXLColumnIndex = pXLColumn[32];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기타-기부금
                    vGDColumnIndex = pGDColumn[33];
                    vXLColumnIndex = pXLColumn[33];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        if (vConvertDecimal != 0)
                        {
                            vConvertString = string.Format("{0}", vConvertDecimal);
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                        }
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //mDONAT_ETC_AMT = iString.ISDecimaltoZero(mDONAT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);
                }
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

        #region ----- XLLINE13_2 -----
        private int XLLINE13_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_SUPPORT_FAMILY, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");
                if (pGridRow == 0)
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 17;
                    //-------------------------------------------------------------------
                    // 다자녀 인원 수
                    vGDColumnIndex = pGDColumn[0];
                    vXLColumnIndex = pXLColumn[0];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                }

                //----[ 3 page ]------------------------------------------------------------------------------------------------------

                if (pGridRow == -1)
                {
                    //-------------------------------------------------------------------
                    vXLine = 134;
                    //-------------------------------------------------------------------

                    // 기본공제
                    vXLColumnIndex = pXLColumn[31];
                    if (iString.ISDecimaltoZero(mBASE_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBASE_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, mBASE_COUNT);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 경로우대
                    vXLColumnIndex = pXLColumn[32];
                    if (iString.ISDecimaltoZero(mOLD_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mOLD_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 출산/입양양육
                    vXLColumnIndex = pXLColumn[33];
                    if (iString.ISDecimaltoZero(mBIRTH_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBIRTH_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-보험료
                    vXLColumnIndex = pXLColumn[37];
                    if (iString.ISDecimaltoZero(mINSURE_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mINSURE_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-의료비
                    vXLColumnIndex = pXLColumn[38];
                    if (iString.ISDecimaltoZero(mMEDICAL_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mMEDICAL_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-교육비
                    vXLColumnIndex = pXLColumn[39];
                    if (iString.ISDecimaltoZero(mEDU_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mEDU_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-신용카드
                    vXLColumnIndex = pXLColumn[40];
                    if (iString.ISDecimaltoZero(mCREDIT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCREDIT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-직불카드
                    vXLColumnIndex = pXLColumn[41];
                    if (iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCHECK_CREDIT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-학원비지로납부액
                    vXLColumnIndex = pXLColumn[42];
                    if (iString.ISDecimaltoZero(mCASH_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCASH_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-현금영수증
                    vXLColumnIndex = pXLColumn[43];
                    if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDONAT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-전통시장사용액 
                    vXLColumnIndex = pXLColumn[44];
                    if (iString.ISDecimaltoZero(mTRAD_MARKET_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mTRAD_MARKET_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-기부금
                    vXLColumnIndex = pXLColumn[45];
                    if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDONAT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }


                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 부녀자
                    vXLColumnIndex = pXLColumn[34];
                    if (iString.ISDecimaltoZero(mBASE_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBASE_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, mBASE_COUNT);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 장애인
                    vXLColumnIndex = pXLColumn[35];
                    if (iString.ISDecimaltoZero(mOLD_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mOLD_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 6세이하
                    vXLColumnIndex = pXLColumn[36];
                    if (iString.ISDecimaltoZero(mBIRTH_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBIRTH_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-보험료
                    vXLColumnIndex = pXLColumn[48];
                    if (iString.ISDecimaltoZero(mINSURE_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mINSURE_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-의료비
                    vXLColumnIndex = pXLColumn[49];
                    if (iString.ISDecimaltoZero(mMEDICAL_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mMEDICAL_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-교육비
                    vXLColumnIndex = pXLColumn[50];
                    if (iString.ISDecimaltoZero(mEDU_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mEDU_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-신용카드
                    vXLColumnIndex = pXLColumn[51];
                    if (iString.ISDecimaltoZero(mCREDIT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCREDIT_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-직불카드
                    vXLColumnIndex = pXLColumn[52];
                    if (iString.ISDecimaltoZero(mCHECK_CREDIT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCHECK_CREDIT_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-학원비지로납부액
                    vXLColumnIndex = pXLColumn[53];
                    if (iString.ISDecimaltoZero(mACADE_GIRO_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mACADE_GIRO_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-현금영수증
                    //vXLColumnIndex = pXLColumn[54];
                    //if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mDONAT_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    // 기타-전통시장사용액 
                    vXLColumnIndex = pXLColumn[55];
                    if (iString.ISDecimaltoZero(mTRAD_MARKET_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mTRAD_MARKET_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-기부금
                    vXLColumnIndex = pXLColumn[56];
                    if (iString.ISDecimaltoZero(mDONAT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDONAT_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }



                    vXLine = pXLine;
                }
                else
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 관계코드
                    vGDColumnIndex = pGDColumn[1];
                    vXLColumnIndex = pXLColumn[1];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 성명
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기본공제
                    vGDColumnIndex = pGDColumn[3];
                    vXLColumnIndex = pXLColumn[3];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 경로우대
                    vGDColumnIndex = pGDColumn[4];
                    vXLColumnIndex = pXLColumn[4];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 출산/입양양육
                    vGDColumnIndex = pGDColumn[5];
                    vXLColumnIndex = pXLColumn[5];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 국세청-보험료
                    vGDColumnIndex = pGDColumn[6];
                    vXLColumnIndex = pXLColumn[6];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mINSURE_AMT = iString.ISDecimaltoZero(mINSURE_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-의료비
                    vGDColumnIndex = pGDColumn[7];
                    vXLColumnIndex = pXLColumn[7];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mMEDICAL_AMT = iString.ISDecimaltoZero(mMEDICAL_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-교육비
                    vGDColumnIndex = pGDColumn[8];
                    vXLColumnIndex = pXLColumn[8];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mEDU_AMT = iString.ISDecimaltoZero(mEDU_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-신용카드
                    vGDColumnIndex = pGDColumn[9];
                    vXLColumnIndex = pXLColumn[9];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mCREDIT_AMT = iString.ISDecimaltoZero(mCREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-직불카드
                    vGDColumnIndex = pGDColumn[10];
                    vXLColumnIndex = pXLColumn[10];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-현금
                    vGDColumnIndex = pGDColumn[11];
                    vXLColumnIndex = pXLColumn[11];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-전통시장
                    vGDColumnIndex = pGDColumn[12];
                    vXLColumnIndex = pXLColumn[12];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-대중교통
                    vGDColumnIndex = pGDColumn[13];
                    vXLColumnIndex = pXLColumn[13];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-기부금
                    vGDColumnIndex = pGDColumn[14];
                    vXLColumnIndex = pXLColumn[14];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mDONAT_AMT = iString.ISDecimaltoZero(mDONAT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 국가타입
                    vGDColumnIndex = pGDColumn[15];
                    vXLColumnIndex = pXLColumn[15];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 주민번호
                    vGDColumnIndex = pGDColumn[16];
                    vXLColumnIndex = pXLColumn[16];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 부녀자
                    vGDColumnIndex = pGDColumn[17];
                    vXLColumnIndex = pXLColumn[17];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mWOMAN_COUNT = iString.ISDecimaltoZero(mWOMAN_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[30]), 0);

                    // 한부모
                    vGDColumnIndex = pGDColumn[18];
                    vXLColumnIndex = pXLColumn[18];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mWOMAN_COUNT = iString.ISDecimaltoZero(mWOMAN_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[30]), 0);

                    // 장애인
                    vGDColumnIndex = pGDColumn[19];
                    vXLColumnIndex = pXLColumn[19];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    // mDISABILITY_COUNT = iString.ISDecimaltoZero(mDISABILITY_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[28]), 0);

                    // 자녀양육(6세이하)
                    vGDColumnIndex = pGDColumn[20];
                    vXLColumnIndex = pXLColumn[20];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHILD_COUNT = iString.ISDecimaltoZero(mCHILD_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[29]), 0);

                    // 기타-보험료
                    vGDColumnIndex = pGDColumn[21];
                    vXLColumnIndex = pXLColumn[21];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mINSURE_ETC_AMT = iString.ISDecimaltoZero(mINSURE_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-의료비
                    vGDColumnIndex = pGDColumn[22];
                    vXLColumnIndex = pXLColumn[22];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mMEDICAL_ETC_AMT = iString.ISDecimaltoZero(mMEDICAL_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-교육비
                    vGDColumnIndex = pGDColumn[23];
                    vXLColumnIndex = pXLColumn[23];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mEDU_ETC_AMT = iString.ISDecimaltoZero(mEDU_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-신용카드
                    vGDColumnIndex = pGDColumn[24];
                    vXLColumnIndex = pXLColumn[24];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCREDIT_ETC_AMT = iString.ISDecimaltoZero(mCREDIT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-직불카드
                    vGDColumnIndex = pGDColumn[25];
                    vXLColumnIndex = pXLColumn[25];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mCHECK_CREDIT_ETC_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-현금영수증
                    vGDColumnIndex = pGDColumn[26];
                    vXLColumnIndex = pXLColumn[26];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mACADE_GIRO_ETC_AMT = iString.ISDecimaltoZero(mACADE_GIRO_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-전통시장
                    vGDColumnIndex = pGDColumn[27];
                    vXLColumnIndex = pXLColumn[27];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기타-대중교통
                    vGDColumnIndex = pGDColumn[28];
                    vXLColumnIndex = pXLColumn[28];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mTRAD_MARKET_ETC_AMT = iString.ISDecimaltoZero(mTRAD_MARKET_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-기부금
                    vGDColumnIndex = pGDColumn[29];
                    vXLColumnIndex = pXLColumn[29];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    //mDONAT_ETC_AMT = iString.ISDecimaltoZero(mDONAT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);
                }
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

        #region ----- XLLINE12_2 -----
        private int XLLINE12_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_SUPPORT_FAMILY, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");
                if (pGridRow == 0)
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 10;
                    //-------------------------------------------------------------------
                    // 다자녀 인원 수
                    vGDColumnIndex = pGDColumn[0];
                    vXLColumnIndex = pXLColumn[0];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                }

                //----[ 3 page ]------------------------------------------------------------------------------------------------------

                if (pGridRow == -1)
                {
                    //-------------------------------------------------------------------
                    vXLine = 131;
                    //-------------------------------------------------------------------

                    // 기본공제
                    vXLColumnIndex = pXLColumn[31];
                    if (iString.ISDecimaltoZero(mBASE_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBASE_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, mBASE_COUNT);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 경로우대
                    vXLColumnIndex = pXLColumn[32];
                    if (iString.ISDecimaltoZero(mOLD_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mOLD_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 출산/입양양육
                    vXLColumnIndex = pXLColumn[33];
                    if (iString.ISDecimaltoZero(mBIRTH_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBIRTH_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-보험료
                    vXLColumnIndex = pXLColumn[37];
                    if (iString.ISDecimaltoZero(mINSURE_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mINSURE_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-의료비
                    vXLColumnIndex = pXLColumn[38];
                    if (iString.ISDecimaltoZero(mMEDICAL_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mMEDICAL_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-교육비
                    vXLColumnIndex = pXLColumn[39];
                    if (iString.ISDecimaltoZero(mEDU_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mEDU_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-신용카드
                    vXLColumnIndex = pXLColumn[40];
                    if (iString.ISDecimaltoZero(mCREDIT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCREDIT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-직불카드
                    vXLColumnIndex = pXLColumn[41];
                    if (iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCHECK_CREDIT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-학원비지로납부액
                    vXLColumnIndex = pXLColumn[42];
                    if (iString.ISDecimaltoZero(mCASH_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCASH_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-현금영수증
                    vXLColumnIndex = pXLColumn[43];
                    if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDONAT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-전통시장사용액 
                    vXLColumnIndex = pXLColumn[44];
                    if (iString.ISDecimaltoZero(mTRAD_MARKET_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mTRAD_MARKET_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-기부금
                    vXLColumnIndex = pXLColumn[45];
                    if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDONAT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }


                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 부녀자
                    vXLColumnIndex = pXLColumn[34];
                    if (iString.ISDecimaltoZero(mBASE_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBASE_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, mBASE_COUNT);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 장애인
                    vXLColumnIndex = pXLColumn[35];
                    if (iString.ISDecimaltoZero(mOLD_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mOLD_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 6세이하
                    vXLColumnIndex = pXLColumn[36];
                    if (iString.ISDecimaltoZero(mBIRTH_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBIRTH_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-보험료
                    vXLColumnIndex = pXLColumn[48];
                    if (iString.ISDecimaltoZero(mINSURE_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mINSURE_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-의료비
                    vXLColumnIndex = pXLColumn[49];
                    if (iString.ISDecimaltoZero(mMEDICAL_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mMEDICAL_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-교육비
                    vXLColumnIndex = pXLColumn[50];
                    if (iString.ISDecimaltoZero(mEDU_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mEDU_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-신용카드
                    vXLColumnIndex = pXLColumn[51];
                    if (iString.ISDecimaltoZero(mCREDIT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCREDIT_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-직불카드
                    vXLColumnIndex = pXLColumn[52];
                    if (iString.ISDecimaltoZero(mCHECK_CREDIT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCHECK_CREDIT_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-학원비지로납부액
                    vXLColumnIndex = pXLColumn[53];
                    if (iString.ISDecimaltoZero(mACADE_GIRO_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mACADE_GIRO_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-현금영수증
                    //vXLColumnIndex = pXLColumn[54];
                    //if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    //{
                    //    vConvertString = string.Format("{0}", mDONAT_AMT);
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}
                    //else
                    //{
                    //    vConvertString = string.Empty;
                    //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    //}

                    // 기타-전통시장사용액 
                    vXLColumnIndex = pXLColumn[55];
                    if (iString.ISDecimaltoZero(mTRAD_MARKET_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mTRAD_MARKET_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 기타-기부금
                    vXLColumnIndex = pXLColumn[56];
                    if (iString.ISDecimaltoZero(mDONAT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDONAT_ETC_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }



                    vXLine = pXLine;
                }
                else
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 관계코드
                    vGDColumnIndex = pGDColumn[1];
                    vXLColumnIndex = pXLColumn[1];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 성명
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기본공제
                    vGDColumnIndex = pGDColumn[31];
                    vXLColumnIndex = pXLColumn[31];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mBASE_COUNT = iString.ISDecimaltoZero(mBASE_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[25]), 0);

                    // 경로우대
                    vGDColumnIndex = pGDColumn[32];
                    vXLColumnIndex = pXLColumn[32];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mOLD_COUNT = iString.ISDecimaltoZero(mOLD_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[26]), 0);

                    // 출산/입양양육
                    vGDColumnIndex = pGDColumn[33];
                    vXLColumnIndex = pXLColumn[33];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mBIRTH_COUNT = iString.ISDecimaltoZero(mBIRTH_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[27]), 0);

                    // 국세청-보험료
                    vGDColumnIndex = pGDColumn[37];
                    vXLColumnIndex = pXLColumn[37];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mINSURE_AMT = iString.ISDecimaltoZero(mINSURE_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-의료비
                    vGDColumnIndex = pGDColumn[38];
                    vXLColumnIndex = pXLColumn[38];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mMEDICAL_AMT = iString.ISDecimaltoZero(mMEDICAL_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-교육비
                    vGDColumnIndex = pGDColumn[39];
                    vXLColumnIndex = pXLColumn[39];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mEDU_AMT = iString.ISDecimaltoZero(mEDU_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-신용카드
                    vGDColumnIndex = pGDColumn[40];
                    vXLColumnIndex = pXLColumn[40];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCREDIT_AMT = iString.ISDecimaltoZero(mCREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-직불카드
                    vGDColumnIndex = pGDColumn[41];
                    vXLColumnIndex = pXLColumn[41];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-학원비지로납부액  
                    vGDColumnIndex = pGDColumn[42];
                    vXLColumnIndex = pXLColumn[42];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mACADE_GIRO_AMT = iString.ISDecimaltoZero(mACADE_GIRO_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-현금영수증
                    vGDColumnIndex = pGDColumn[43];
                    vXLColumnIndex = pXLColumn[43];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCASH_AMT = iString.ISDecimaltoZero(mCASH_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-전통시장사용액
                    vGDColumnIndex = pGDColumn[44];
                    vXLColumnIndex = pXLColumn[44];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mTRAD_MARKET_AMT = iString.ISDecimaltoZero(mTRAD_MARKET_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-기부금
                    vGDColumnIndex = pGDColumn[45];
                    vXLColumnIndex = pXLColumn[45];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mDONAT_AMT = iString.ISDecimaltoZero(mDONAT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 국가타입
                    vGDColumnIndex = pGDColumn[46];
                    vXLColumnIndex = pXLColumn[46];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 주민번호
                    vGDColumnIndex = pGDColumn[47];
                    vXLColumnIndex = pXLColumn[47];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 부녀자
                    vGDColumnIndex = pGDColumn[34];
                    vXLColumnIndex = pXLColumn[34];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mWOMAN_COUNT = iString.ISDecimaltoZero(mWOMAN_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[30]), 0);

                    // 장애인
                    vGDColumnIndex = pGDColumn[35];
                    vXLColumnIndex = pXLColumn[35];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mDISABILITY_COUNT = iString.ISDecimaltoZero(mDISABILITY_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[28]), 0);

                    // 자녀양육(6세이하)
                    vGDColumnIndex = pGDColumn[36];
                    vXLColumnIndex = pXLColumn[36];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCHILD_COUNT = iString.ISDecimaltoZero(mCHILD_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[29]), 0);

                    // 기타-보험료
                    vGDColumnIndex = pGDColumn[48];
                    vXLColumnIndex = pXLColumn[48];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mINSURE_ETC_AMT = iString.ISDecimaltoZero(mINSURE_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-의료비
                    vGDColumnIndex = pGDColumn[49];
                    vXLColumnIndex = pXLColumn[49];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mMEDICAL_ETC_AMT = iString.ISDecimaltoZero(mMEDICAL_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-교육비
                    vGDColumnIndex = pGDColumn[50];
                    vXLColumnIndex = pXLColumn[50];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mEDU_ETC_AMT = iString.ISDecimaltoZero(mEDU_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-신용카드
                    vGDColumnIndex = pGDColumn[51];
                    vXLColumnIndex = pXLColumn[51];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCREDIT_ETC_AMT = iString.ISDecimaltoZero(mCREDIT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-직불카드
                    vGDColumnIndex = pGDColumn[52];
                    vXLColumnIndex = pXLColumn[52];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCHECK_CREDIT_ETC_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-학원비지로납부액
                    vGDColumnIndex = pGDColumn[53];
                    vXLColumnIndex = pXLColumn[53];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mACADE_GIRO_ETC_AMT = iString.ISDecimaltoZero(mACADE_GIRO_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-현금
                    vGDColumnIndex = pGDColumn[54];
                    vXLColumnIndex = pXLColumn[54];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기타-전통시장사용액
                    vGDColumnIndex = pGDColumn[55];
                    vXLColumnIndex = pXLColumn[55];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mTRAD_MARKET_ETC_AMT = iString.ISDecimaltoZero(mTRAD_MARKET_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-기부금
                    vGDColumnIndex = pGDColumn[56];
                    vXLColumnIndex = pXLColumn[56];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mDONAT_ETC_AMT = iString.ISDecimaltoZero(mDONAT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);
                }
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

        #region ----- XLLINE11_2 -----

        private int XLLine11_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_SUPPORT_FAMILY, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");
                if (pGridRow == 0)
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 10;
                    //-------------------------------------------------------------------
                    // 다자녀 인원 수
                    vGDColumnIndex = pGDColumn[0];
                    vXLColumnIndex = pXLColumn[0];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                }

                //----[ 3 page ]------------------------------------------------------------------------------------------------------

                if (pGridRow == -1)
                {
                    //-------------------------------------------------------------------
                    vXLine = 131;
                    //-------------------------------------------------------------------

                    // 기본공제
                    vXLColumnIndex = pXLColumn[25];
                    if (iString.ISDecimaltoZero(mBASE_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBASE_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, mBASE_COUNT);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 경로우대
                    vXLColumnIndex = pXLColumn[26];
                    if (iString.ISDecimaltoZero(mOLD_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mOLD_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 출산/입양양육
                    vXLColumnIndex = pXLColumn[27];
                    if (iString.ISDecimaltoZero(mBIRTH_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mBIRTH_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 장애인
                    vXLColumnIndex = pXLColumn[28];
                    if (iString.ISDecimaltoZero(mDISABILITY_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDISABILITY_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 자녀양육(6세이하)
                    vXLColumnIndex = pXLColumn[29];
                    if (iString.ISDecimaltoZero(mCHILD_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCHILD_COUNT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-보험료
                    vXLColumnIndex = pXLColumn[8];
                    if (iString.ISDecimaltoZero(mINSURE_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mINSURE_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-의료비
                    vXLColumnIndex = pXLColumn[9];
                    if (iString.ISDecimaltoZero(mMEDICAL_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mMEDICAL_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-교육비
                    vXLColumnIndex = pXLColumn[10];
                    if (iString.ISDecimaltoZero(mEDU_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mEDU_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-신용카드
                    vXLColumnIndex = pXLColumn[11];
                    if (iString.ISDecimaltoZero(mCREDIT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCREDIT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-직불카드
                    vXLColumnIndex = pXLColumn[12];
                    if (iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCHECK_CREDIT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-현금
                    vXLColumnIndex = pXLColumn[13];
                    if (iString.ISDecimaltoZero(mCASH_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCASH_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-기부금
                    vXLColumnIndex = pXLColumn[14];
                    if (iString.ISDecimaltoZero(mDONAT_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDONAT_AMT);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 부녀자
                    vXLColumnIndex = pXLColumn[30];
                    if (iString.ISDecimaltoZero(mWOMAN_COUNT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mWOMAN_COUNT);
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-보험료
                    vXLColumnIndex = pXLColumn[8];
                    if (iString.ISDecimaltoZero(mINSURE_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mINSURE_ETC_AMT);
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-의료비
                    vXLColumnIndex = pXLColumn[9];
                    if (iString.ISDecimaltoZero(mMEDICAL_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mMEDICAL_ETC_AMT);
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-교육비
                    vXLColumnIndex = pXLColumn[10];
                    if (iString.ISDecimaltoZero(mEDU_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mEDU_ETC_AMT);
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-신용카드
                    vXLColumnIndex = pXLColumn[11];
                    if (iString.ISDecimaltoZero(mCREDIT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCREDIT_ETC_AMT);
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-직불카드
                    vXLColumnIndex = pXLColumn[12];
                    if (iString.ISDecimaltoZero(mCHECK_CREDIT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mCHECK_CREDIT_AMT);
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }

                    // 국세청-기부금
                    vXLColumnIndex = pXLColumn[14];
                    if (iString.ISDecimaltoZero(mDONAT_ETC_AMT, 0) != 0)
                    {
                        vConvertString = string.Format("{0}", mDONAT_ETC_AMT);
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString);
                    }

                    vXLine = pXLine;
                }
                else
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 관계코드
                    vGDColumnIndex = pGDColumn[1];
                    vXLColumnIndex = pXLColumn[1];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 성명
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기본공제
                    vGDColumnIndex = pGDColumn[3];
                    vXLColumnIndex = pXLColumn[3];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mBASE_COUNT = iString.ISDecimaltoZero(mBASE_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[25]), 0);

                    // 경로우대
                    vGDColumnIndex = pGDColumn[4];
                    vXLColumnIndex = pXLColumn[4];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mOLD_COUNT = iString.ISDecimaltoZero(mOLD_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[26]), 0);

                    // 출산/입양양육
                    vGDColumnIndex = pGDColumn[5];
                    vXLColumnIndex = pXLColumn[5];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mBIRTH_COUNT = iString.ISDecimaltoZero(mBIRTH_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[27]), 0);

                    // 장애인
                    vGDColumnIndex = pGDColumn[6];
                    vXLColumnIndex = pXLColumn[6];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mDISABILITY_COUNT = iString.ISDecimaltoZero(mDISABILITY_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[28]), 0);

                    // 자녀양육(6세이하)
                    vGDColumnIndex = pGDColumn[7];
                    vXLColumnIndex = pXLColumn[7];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCHILD_COUNT = iString.ISDecimaltoZero(mCHILD_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[29]), 0);

                    // 국세청-보험료
                    vGDColumnIndex = pGDColumn[8];
                    vXLColumnIndex = pXLColumn[8];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mINSURE_AMT = iString.ISDecimaltoZero(mINSURE_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-의료비
                    vGDColumnIndex = pGDColumn[9];
                    vXLColumnIndex = pXLColumn[9];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mMEDICAL_AMT = iString.ISDecimaltoZero(mMEDICAL_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-교육비
                    vGDColumnIndex = pGDColumn[10];
                    vXLColumnIndex = pXLColumn[10];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mEDU_AMT = iString.ISDecimaltoZero(mEDU_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-신용카드
                    vGDColumnIndex = pGDColumn[11];
                    vXLColumnIndex = pXLColumn[11];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCREDIT_AMT = iString.ISDecimaltoZero(mCREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-직불카드
                    vGDColumnIndex = pGDColumn[12];
                    vXLColumnIndex = pXLColumn[12];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCHECK_CREDIT_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-현금
                    vGDColumnIndex = pGDColumn[13];
                    vXLColumnIndex = pXLColumn[13];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCASH_AMT = iString.ISDecimaltoZero(mCASH_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 국세청-기부금
                    vGDColumnIndex = pGDColumn[14];
                    vXLColumnIndex = pXLColumn[14];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mDONAT_AMT = iString.ISDecimaltoZero(mDONAT_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    // 국가타입
                    vGDColumnIndex = pGDColumn[15];
                    vXLColumnIndex = pXLColumn[15];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 주민번호
                    vGDColumnIndex = pGDColumn[16];
                    vXLColumnIndex = pXLColumn[16];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 부녀자
                    vGDColumnIndex = pGDColumn[17];
                    vXLColumnIndex = pXLColumn[17];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mWOMAN_COUNT = iString.ISDecimaltoZero(mWOMAN_COUNT, 0) + iString.ISDecimaltoZero(pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, pGDColumn[30]), 0);

                    // 기타-보험료
                    vGDColumnIndex = pGDColumn[18];
                    vXLColumnIndex = pXLColumn[18];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mINSURE_ETC_AMT = iString.ISDecimaltoZero(mINSURE_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-의료비
                    vGDColumnIndex = pGDColumn[19];
                    vXLColumnIndex = pXLColumn[19];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mMEDICAL_ETC_AMT = iString.ISDecimaltoZero(mMEDICAL_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-교육비
                    vGDColumnIndex = pGDColumn[20];
                    vXLColumnIndex = pXLColumn[20];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mEDU_ETC_AMT = iString.ISDecimaltoZero(mEDU_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-신용카드
                    vGDColumnIndex = pGDColumn[21];
                    vXLColumnIndex = pXLColumn[21];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCREDIT_ETC_AMT = iString.ISDecimaltoZero(mCREDIT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-직불카드
                    vGDColumnIndex = pGDColumn[22];
                    vXLColumnIndex = pXLColumn[22];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mCHECK_CREDIT_ETC_AMT = iString.ISDecimaltoZero(mCHECK_CREDIT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);

                    // 기타-현금
                    vGDColumnIndex = pGDColumn[23];
                    vXLColumnIndex = pXLColumn[23];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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

                    // 기타-기부금
                    vGDColumnIndex = pGDColumn[24];
                    vXLColumnIndex = pXLColumn[24];
                    vObject = pGrid_SUPPORT_FAMILY.GetCellValue(pGridRow, vGDColumnIndex);
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
                    mDONAT_ETC_AMT = iString.ISDecimaltoZero(mDONAT_ETC_AMT, 0) + iString.ISDecimaltoZero(vObject, 0);
                }
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

        #region ----- XLHeader_PRINT_SAVING_INFO(2011, 2012) -----

        private int XLHeader_PRINT_SAVING_INFO(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호
            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab2");


                // 법인명(상호)
                vXLine = 6;
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = 8;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 사업자등록번호
                vXLine = 6;
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = 29;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 성명
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = 8;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주민번호
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = 29;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주소
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = 8;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 개인 전화번호
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TELEPHON_NO");
                vXLColumnIndex = 36;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업장 소재지
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = 8;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 사업장 전화번호
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TEL_NUMBER");
                vXLColumnIndex = 36;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
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

        #region ----- XLHeader_PRINT_SAVING_INFO(2013) -----

        private int XLHeader_PRINT_SAVING_INFO_2013(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호
            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab2");

                // 법인명(상호)
                vXLine = 4;
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = 15;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 사업자등록번호
                vXLine = 4;
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = 29;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 성명
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = 15;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주민번호
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = 29;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주소
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = 15;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 개인 전화번호
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TELEPHON_NO");
                vXLColumnIndex = 36;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업장 소재지
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = 15;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업장 전화번호
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TEL_NUMBER");
                vXLColumnIndex = 36;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
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

        #region ----- XLHeader_PRINT_HOUSE_INFO(2013) -----

        private int XLHeader_PRINT_HOUSE_INFO_2013(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호
            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab3");

                // 법인명(상호)
                vXLine = 4;
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = 15;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 사업자등록번호
                vXLine = 4;
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = 29;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 성명
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = 15;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 주민번호
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = 29;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주소
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = 15;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 개인 전화번호
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TELEPHON_NO");
                vXLColumnIndex = 36;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업장 소재지
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = 15;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 사업장 전화번호
                vGDColumnIndex = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TEL_NUMBER");
                vXLColumnIndex = 36;
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);
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
                vXLine = vXLine + 1;
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

        #region ----- XLLINE3 연금/저축소득공제-----

        private int XLLine3(InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab2");

                // 저축TYPE명[저축구분]
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 금융기관명
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 계좌번호
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 불입금액
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 공제금액
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
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

        #region ----- XLLINE4 연금/저축소득공제-----

        private int XLLine4(InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab2");

                // 금융기관명
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 계좌번호
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
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

                // 납입연차
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 불입금액
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 공제금액
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGRID.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
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

        #region ----- 월세액 소득공제 명세 (2013)-----

        private int XLLine5(InfoSummit.Win.ControlAdv.ISDataAdapter pidaHOUSE_LEASE_INFO_10, int pXLine)
        {
            string vMessage = string.Empty;

            string[] vDBColumn;
            int[] vXLColumn;

            int vPrintingLine = pXLine;

            try
            {
                int vTotalRow = pidaHOUSE_LEASE_INFO_10.CurrentRows.Count;

                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    SetArray_House1(out vDBColumn, out vXLColumn);

                    //mPrinting.XLSetCell(2, 6, "●");

                    foreach (System.Data.DataRow vRow in pidaHOUSE_LEASE_INFO_10.CurrentRows)
                    {
                        vCountRow++;

                        vPrintingLine = XlLine_House1(vRow, vPrintingLine, vDBColumn, vXLColumn);


                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return pXLine;
        }
        #endregion;

        #region ----- 거주자간 주택임차차입금 원리금 상환액 소득공제 명세(2013)-----

        private int XLLine6(InfoSummit.Win.ControlAdv.ISDataAdapter pidaHOUSE_LEASE_INFO_20, int pXLine)
        {
            string vMessage = string.Empty;

            string[] vDBColumn;
            int[] vXLColumn;

            int vPrintingLine = pXLine;

            try
            {
                int vTotalRow = pidaHOUSE_LEASE_INFO_20.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    SetArray_House2(out vDBColumn, out vXLColumn);

                    foreach (System.Data.DataRow vRow in pidaHOUSE_LEASE_INFO_20.CurrentRows)
                    {
                        vCountRow++;

                        vPrintingLine = XlLine_House2(vRow, vPrintingLine, vDBColumn, vXLColumn);

                        vPrintingLine = vPrintingLine + 12;

                        vPrintingLine = XlLine_House3(vRow, vPrintingLine, vDBColumn, vXLColumn);

                        vPrintingLine = vPrintingLine - 14;


                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return pXLine;
        }


        #endregion;

        #region ----- 월세액 소득공제 명세 (2014)-----

        private int XLLine7(InfoSummit.Win.ControlAdv.ISDataAdapter pidaHOUSE_LEASE_INFO_10, int pXLine)
        {
            string vMessage = string.Empty;

            string[] vDBColumn;
            int[] vXLColumn;

            int vPrintingLine = pXLine;

            try
            {
                int vTotalRow = pidaHOUSE_LEASE_INFO_10.CurrentRows.Count;

                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    SetArray_House3(out vDBColumn, out vXLColumn);

                    //mPrinting.XLSetCell(2, 6, "●");

                    foreach (System.Data.DataRow vRow in pidaHOUSE_LEASE_INFO_10.CurrentRows)
                    {
                        vCountRow++;

                        vPrintingLine = XlLine_House4(vRow, vPrintingLine, vDBColumn, vXLColumn);


                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return pXLine;
        }
        #endregion;

        #region ----- 거주자간 주택임차차입금 원리금 상환액 소득공제 명세(2017)-----

        private int XLLine8(InfoSummit.Win.ControlAdv.ISDataAdapter pidaHOUSE_LEASE_INFO_20, int pXLine)
        {
            string vMessage = string.Empty;

            string[] vDBColumn;
            int[] vXLColumn;

            int vPrintingLine = pXLine;

            try
            {
                int vTotalRow = pidaHOUSE_LEASE_INFO_20.CurrentRows.Count;
                if (vTotalRow > 0)
                {
                    int vCountRow = 0;

                    SetArray_House4(out vDBColumn, out vXLColumn);

                    foreach (System.Data.DataRow vRow in pidaHOUSE_LEASE_INFO_20.CurrentRows)
                    {
                        vCountRow++;

                        vPrintingLine = XlLine_House2(vRow, vPrintingLine, vDBColumn, vXLColumn);

                        vPrintingLine = vPrintingLine + 12;

                        vPrintingLine = XlLine_House5(vRow, vPrintingLine, vDBColumn, vXLColumn);

                        vPrintingLine = vPrintingLine - 14;


                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return pXLine;
        }


        #endregion;

        #region ----- XlLine_House1 -----

        private int XlLine_House1(System.Data.DataRow pRow, int pPrintingLine, string[] pDBColumn, int[] pXLColumn)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            string vColumnName1 = string.Empty;

            int vXLColumnIndex = 0;

            string vConvertString1 = string.Empty;

            System.DateTime vConvertDateTime = new System.DateTime();

            decimal vConvertDecimal = 0m;

            bool IsConvert1 = false;

            try
            {

                //[임대인 성명]
                vColumnName1 = pDBColumn[3];
                vXLColumnIndex = pXLColumn[3];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                    mPrinting.XLSetCell(2, 6, "●");
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주민등록번호]
                vColumnName1 = pDBColumn[4];
                vXLColumnIndex = pXLColumn[4];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[임대차계약서 상 주소지]
                vColumnName1 = pDBColumn[6];
                vXLColumnIndex = pXLColumn[6];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[임대차계약기간]
                vColumnName1 = pDBColumn[8];
                vXLColumnIndex = pXLColumn[8];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}~", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[임대차계약기간 to]
                vColumnName1 = pDBColumn[9];
                vXLColumnIndex = pXLColumn[9];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                }


                //[월세액]
                vColumnName1 = pDBColumn[10];
                vXLColumnIndex = pXLColumn[10];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[공제금액]
                vColumnName1 = pDBColumn[11];
                vXLColumnIndex = pXLColumn[11];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
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

        #region ----- XlLine_House2 -----

        private int XlLine_House2(System.Data.DataRow pRow, int pPrintingLine, string[] pDBColumn, int[] pXLColumn)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            string vColumnName1 = string.Empty;

            int vXLColumnIndex = 0;

            string vConvertString1 = string.Empty;

            System.DateTime vConvertDateTime = new System.DateTime();

            decimal vConvertDecimal = 0m;

            bool IsConvert1 = false;

            try
            {
                //[대주]
                vColumnName1 = pDBColumn[3];
                vXLColumnIndex = pXLColumn[3];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                    mPrinting.XLSetCell(2, 13, "●");
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주민등록번호]
                vColumnName1 = pDBColumn[4];
                vXLColumnIndex = pXLColumn[4];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[임대차계약기간]
                vColumnName1 = pDBColumn[5];
                vXLColumnIndex = pXLColumn[5];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}~", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[임대차계약기간 to]
                vColumnName1 = pDBColumn[6];
                vXLColumnIndex = pXLColumn[6];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 1, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                }


                //[차입금 이자율]
                vColumnName1 = pDBColumn[7];
                vXLColumnIndex = pXLColumn[7];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###.###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[원리금 상환액 - 계]
                vColumnName1 = pDBColumn[8];
                vXLColumnIndex = pXLColumn[8];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[원리금상환액 - 원리금]
                vColumnName1 = pDBColumn[9];
                vXLColumnIndex = pXLColumn[9];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[원리금상환액 - 이자]
                vColumnName1 = pDBColumn[10];
                vXLColumnIndex = pXLColumn[10];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[공제금액]
                vColumnName1 = pDBColumn[11];
                vXLColumnIndex = pXLColumn[11];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
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

        #region ----- XlLine_House3 -----

        private int XlLine_House3(System.Data.DataRow pRow, int pPrintingLine, string[] pDBColumn, int[] pXLColumn)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            string vColumnName1 = string.Empty;
            string vColumnName2 = string.Empty;

            int vXLColumnIndex = 0;

            string vConvertString1 = string.Empty;

            System.DateTime vConvertDateTime = new System.DateTime();
            System.DateTime vConvertDateTime2 = new System.DateTime();

            decimal vConvertDecimal = 0m;

            bool IsConvert1 = false;

            try
            {
                //[임대인성명]
                vColumnName1 = pDBColumn[12];
                vXLColumnIndex = pXLColumn[12];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주민등록번호]
                vColumnName1 = pDBColumn[13];
                vXLColumnIndex = pXLColumn[13];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주소지]
                vColumnName1 = pDBColumn[15];
                vXLColumnIndex = pXLColumn[15];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[금전소비대차 계약기간]
                vColumnName1 = pDBColumn[17];
                vColumnName2 = pDBColumn[18];
                vXLColumnIndex = pXLColumn[17];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);
                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime2);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}~{1}", vConvertDateTime.ToShortDateString(), vConvertDateTime2.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[전세보증금
                vColumnName1 = pDBColumn[19];
                vXLColumnIndex = pXLColumn[19];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
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

        #region ----- XlLine_House4 ( 2014년 월세액 소득공제 명세) -----

        private int XlLine_House4(System.Data.DataRow pRow, int pPrintingLine, string[] pDBColumn, int[] pXLColumn)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            string vColumnName1 = string.Empty;

            int vXLColumnIndex = 0;

            string vConvertString1 = string.Empty;

            System.DateTime vConvertDateTime = new System.DateTime();

            decimal vConvertDecimal = 0m;

            bool IsConvert1 = false;

            try
            {

                //[임대인 성명]
                vColumnName1 = pDBColumn[3];
                vXLColumnIndex = pXLColumn[3];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                    mPrinting.XLSetCell(2, 6, "●");
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주민등록번호]
                vColumnName1 = pDBColumn[4];
                vXLColumnIndex = pXLColumn[4];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[임대차계약서 상 주소지]
                vColumnName1 = pDBColumn[6];
                vXLColumnIndex = pXLColumn[6];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[임대차계약기간]
                vColumnName1 = pDBColumn[8];
                vXLColumnIndex = pXLColumn[8];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[임대차계약기간 to]
                vColumnName1 = pDBColumn[9];
                vXLColumnIndex = pXLColumn[9];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);

                }

                //[월세액]
                vColumnName1 = pDBColumn[10];
                vXLColumnIndex = pXLColumn[10];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[공제금액]
                vColumnName1 = pDBColumn[11];
                vXLColumnIndex = pXLColumn[11];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }


                //[주택유형]
                vColumnName1 = pDBColumn[12];
                vXLColumnIndex = pXLColumn[12];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주택계약면적]
                vColumnName1 = pDBColumn[13];
                vXLColumnIndex = pXLColumn[13];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
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

        #region ----- XlLine_House5 ( 2014년 거주지간 주택임차차입금 원리금 상환액 소득공제 명세) -----

        private int XlLine_House5(System.Data.DataRow pRow, int pPrintingLine, string[] pDBColumn, int[] pXLColumn)
        {
            int vXLine = pPrintingLine; //엑셀에 내용이 표시되는 행 번호

            string vColumnName1 = string.Empty;

            int vXLColumnIndex = 0;

            string vConvertString1 = string.Empty;

            System.DateTime vConvertDateTime = new System.DateTime();

            decimal vConvertDecimal = 0m;

            bool IsConvert1 = false;

            try
            {
                //[임대인성명]
                vColumnName1 = pDBColumn[12];
                vXLColumnIndex = pXLColumn[12];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주민등록번호]
                vColumnName1 = pDBColumn[13];
                vXLColumnIndex = pXLColumn[13];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주소지]
                vColumnName1 = pDBColumn[15];
                vXLColumnIndex = pXLColumn[15];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[금전소비대차 계약기간]
                vColumnName1 = pDBColumn[17];
                vXLColumnIndex = pXLColumn[17];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[금전소비대차 계약기간]
                vColumnName1 = pDBColumn[18];
                vXLColumnIndex = pXLColumn[18];

                IsConvert1 = IsConvertDate(pRow[vColumnName1], out vConvertDateTime);

                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }


                //[전세보증금
                vColumnName1 = pDBColumn[19];
                vXLColumnIndex = pXLColumn[19];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주택유형]
                vColumnName1 = pDBColumn[20];
                vXLColumnIndex = pXLColumn[20];
                IsConvert1 = IsConvertString(pRow[vColumnName1], out vConvertString1);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //[주택계약면적]
                vColumnName1 = pDBColumn[21];
                vXLColumnIndex = pXLColumn[21];
                IsConvert1 = IsConvertNumber(pRow[vColumnName1], out vConvertDecimal);
                if (IsConvert1 == true)
                {
                    vConvertString1 = string.Format("{0}", vConvertString1);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertDecimal);
                }
                else
                {
                    vConvertString1 = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString1);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
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

        //30장씩
        #region ----- Excel Main Wirte  Method ----
         
        public int WriteMain(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX_13
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_SUPPORT_FAMILY
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGR_PRINT_SAVING_INFO_2
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGR_PRINT_SAVING_INFO_3
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGR_PRINT_SAVING_INFO_4
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGR_PRINT_SAVING_INFO_5
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGR_PRINT_SAVING_INFO_6
                            , object vPrintDate
                            , string vPrint_Type
                            , object vPrint_Type_Desc
                            , string pOutChoice
                            , decimal pPrintPage
                            , string pPRINT_SAVING_YN
                            , string pPrint_Year
                            , string pPRINT_HOUSE_YN
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pidaHOUSE_LEASE_INFO_10
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pidaHOUSE_LEASE_INFO_20
                           )
        {
            string vMessageText = string.Empty;
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn_1;
            int[] vXLColumn_1;

            int[] vGDColumn_2;
            int[] vXLColumn_2;

            int[] vGDColumn_3;
            int[] vXLColumn_3;

            int[] vGDColumn_4;
            int[] vXLColumn_4;

            int[] vGDColumn_5;
            int[] vXLColumn_5;

            int[] vGDColumn_6;
            int[] vXLColumn_6;

            int[] vGDColumn_7;
            int[] vXLColumn_7;

            int[] vGDColumn_8;
            int[] vXLColumn_8;


            int vTotalRow1 = pGrid_WITHHOLDING_TAX.RowCount;
            int vTotalRow2 = pGrid_SUPPORT_FAMILY.RowCount;
            int vTotalRow3 = pGrid_WITHHOLDING_TAX_13.RowCount;

            int vRowCount = 0;

            int vPrintingLine = 0;

            //int vSecondPrinting = 9; //1인당 3페이지이므로, 3*10=30번째에 인쇄
            int vCountPrinting = 0;


            SetArray1(pGrid_WITHHOLDING_TAX, out vGDColumn_1, out vXLColumn_1);
            SetArray2(pGrid_SUPPORT_FAMILY, out vGDColumn_2, out vXLColumn_2);

            //연금저축 
            SetArray3(pGR_PRINT_SAVING_INFO_2, out vGDColumn_3, out vXLColumn_3);
            SetArray4(pGR_PRINT_SAVING_INFO_5, out vGDColumn_4, out vXLColumn_4);

            //2013전용
            SetArray5(pGrid_WITHHOLDING_TAX_13, out vGDColumn_5, out vXLColumn_5);
            SetArray6(pGrid_SUPPORT_FAMILY, out vGDColumn_6, out vXLColumn_6);

            //2014전용
            SetArray7(pGrid_WITHHOLDING_TAX_13, out vGDColumn_7, out vXLColumn_7);
            SetArray8(pGrid_SUPPORT_FAMILY, out vGDColumn_8, out vXLColumn_8);

            bool isOpen = false;
            for (int vRow1 = 0; vRow1 < vTotalRow1 || vRow1 < vTotalRow3; vRow1++)
            {
                vRowCount++;

                //-------------------------------------------------------------------------------------

                try
                {
                    if (pPrint_Year == "2011")
                    {
                        mPrinting.XLOpenFile("HRMF0705_001.xls");
                        isOpen = true;
                    }
                    else if (pPrint_Year == "2012")
                    {
                        mPrinting.XLOpenFile("HRMF0705_001_12.xls");
                        isOpen = true;
                    }
                    else if (pPrint_Year == "2013")
                    {
                        mPrinting.XLOpenFile("HRMF0705_001_13.xls");
                        isOpen = true;
                    }
                    else //2014
                    {
                        mPrinting.XLOpenFile("HRMF0705_001_14.xls");
                        isOpen = true;
                    }
                }
                catch
                {
                    isOpen = false;
                }

                //-------------------------------------------------------------------------------------

                string vSaveFileName = string.Empty;
                if (pOutChoice == "PDF")
                {
                    System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                    //저장 파일 이름 : 

                    string vName = pGrid_WITHHOLDING_TAX_13.GetCellValue(vRow1, 1).ToString();

                    vSaveFileName = string.Format("{0}_{1}", "근로소득영수증", vName);
                    vSaveFileName = SetExportFileName(vSaveFileName);

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);

                    if (vFileName.Exists)
                    {
                        vFileName.Delete();
                    }

                }
                
                vMessageText = string.Format("{0} - Printing : {1}/{2}", vPrint_Type_Desc, vRowCount, vTotalRow1);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                if (isOpen == true)
                {
                    vCountPrinting++;

                    mCopyLineSUM = CopyAndPaste_1(mCopyLineSUM);
                    vPrintingLine = 1; //(mCopyLineSUM - mIncrementCopyMAX_1) + (mPrintingLineSTART_1 - 1);

                    if (pPrint_Year == "2011" || pPrint_Year == "2012")
                    {
                        pGrid_WITHHOLDING_TAX.CurrentCellMoveTo(vRow1, 0);
                        pGrid_WITHHOLDING_TAX.CurrentCellActivate(vRow1, 0);
                        pGrid_WITHHOLDING_TAX.Focus();
                    }
                    else //2013, 2014
                    {
                        pGrid_WITHHOLDING_TAX_13.CurrentCellMoveTo(vRow1, 0);
                        pGrid_WITHHOLDING_TAX_13.CurrentCellActivate(vRow1, 0);
                        pGrid_WITHHOLDING_TAX_13.Focus();
                    }                    

                    vMessageText = string.Format("{0} - {1}", vMessageText, "근로소득원천징수영수증");
                    mAppInterface.OnAppMessageEvent(vMessageText);

                    // 근로소득원천징수영수증 page 1 - 2.
                    //int vLinePrinting_1 = vPrintingLine + 3;
                    if (pPrint_Year == "2011")
                    {
                        vPrintingLine = XLLine11(pGrid_WITHHOLDING_TAX, vRow1, vPrintingLine, vGDColumn_1, vXLColumn_1, vPrintDate, vPrint_Type_Desc);
                    }
                    else if (pPrint_Year == "2012")
                    {
                        vPrintingLine = XLLine12(pGrid_WITHHOLDING_TAX, vRow1, vPrintingLine, vGDColumn_1, vXLColumn_1, vPrintDate, vPrint_Type_Desc);
                    }
                    else if (pPrint_Year == "2013")
                    {
                        vPrintingLine = XLLine13(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_5, vXLColumn_5, vPrintDate, vPrint_Type, vPrint_Type_Desc);
                    }
                    else //2014
                    {
                        vPrintingLine = XLLine14(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_7, vXLColumn_7, vPrintDate, vPrint_Type, vPrint_Type_Desc);
                    }
                    //---------------------------------------------------------------------------------------------------------------------

                    vMessageText = string.Format("{0} - {1}", vMessageText, "부양가족내역");
                    mAppInterface.OnAppMessageEvent(vMessageText);

                    //부양가족내역 page 3.
                    mINSURE_AMT = 0;    //보험료.
                    mMEDICAL_AMT = 0;   //의료비.
                    mEDU_AMT = 0;       //교육비.
                    mCREDIT_AMT = 0;    //신용카드.
                    mCHECK_CREDIT_AMT = 0;  //체크카드.
                    mACADE_GIRO_AMT = 0;   //학원비지로납부액
                    mCASH_AMT = 0;      //현금영수증.
                    mTRAD_MARKET_AMT = 0;   //전통시장사용액
                    mPUBLIC_TRANSIT_AMT = 0; //대중교통사용액
                    mDONAT_AMT = 0;     //기부금.

                    // 기타.
                    mINSURE_ETC_AMT = 0;    //보험료.
                    mMEDICAL_ETC_AMT = 0;   //의료비.
                    mEDU_ETC_AMT = 0;       //교육비.
                    mCREDIT_ETC_AMT = 0;    //신용카드.
                    mCHECK_CREDIT_ETC_AMT = 0;  //체크카드.
                    mACADE_GIRO_ETC_AMT = 0;   //학원비지로납부액
                    mTRAD_MARKET_ETC_AMT = 0;   //전통시장사용액
                    mETC_PUBLIC_TRANSIT_AMT = 0; //대중교통사용액
                    mDONAT_ETC_AMT = 0;     //기부금.

                    //인원수
                    mBASE_COUNT = 0;    //기본인원수.
                    mOLD_COUNT = 0;     //경로인원수.
                    mBIRTH_COUNT = 0;   //출생인원수.
                    mDISABILITY_COUNT = 0;  //장애인인원수.
                    mCHILD_COUNT = 0;   //6세이하인원수.
                    mWOMAN_COUNT = 0;   //부녀세대

                    int vPrintingLine_2 = vPrintingLine + 8;
                    for (int vRow2 = 0; vRow2 < vTotalRow2; vRow2++)
                    {
                        if (pPrint_Year == "2011")
                        {
                            vPrintingLine = XLLine11_2(pGrid_SUPPORT_FAMILY, vRow2, vPrintingLine, vGDColumn_2, vXLColumn_2);
                        }
                        else if (pPrint_Year == "2012")
                        {
                            vPrintingLine = XLLINE12_2(pGrid_SUPPORT_FAMILY, vRow2, vPrintingLine, vGDColumn_2, vXLColumn_2);
                        }
                        else if (pPrint_Year == "2013")
                        {
                            vPrintingLine = XLLINE13_2(pGrid_SUPPORT_FAMILY, vRow2, vPrintingLine, vGDColumn_6, vXLColumn_6);
                        }
                        else
                        {
                            vPrintingLine = XLLINE14_2(pGrid_SUPPORT_FAMILY, vRow2, vPrintingLine, vGDColumn_8, vXLColumn_8);
                        }
                    }

                    //부양가족 합계 금액 인쇄.
                    //if (pPrint_Year == "2011 ")
                    //{
                    //    vPrintingLine = XLLine2(pGrid_SUPPORT_FAMILY, -1, vPrintingLine, vGDColumn_2, vXLColumn_2);
                    //}
                    //else
                    //{
                    //vPrintingLine = XLLINE2_2(pGrid_SUPPORT_FAMILY, -1, vPrintingLine, vGDColumn_2, vXLColumn_2);
                    //}
                    //---------------------------------------------------------------------------------------------------------------------

                    //연금/저축 등 소득공재 명세서 Page 5

                    if (pPRINT_SAVING_YN == "Y")
                    {
                        //연금/저축 등 소득공재 명세서 출력할 것이 있을 경우만 
                        if (pGR_PRINT_SAVING_INFO_2.RowCount > 0 || pGR_PRINT_SAVING_INFO_3.RowCount > 0 || pGR_PRINT_SAVING_INFO_4.RowCount > 0 || pGR_PRINT_SAVING_INFO_5.RowCount > 0)
                        {
                            mCopyLineSUM++;

                            //있을경우 PAGE 수는 4로 증가 
                            pPrintPage = 4;

                            if (pPrint_Year == "2011" || pPrint_Year == "2012")
                            {
                                vPrintingLine = 6;
                                vPrintingLine = XLHeader_PRINT_SAVING_INFO(pGrid_WITHHOLDING_TAX, vRow1, vPrintingLine, vGDColumn_1);

                                vMessageText = string.Format("{0} - {1}", vMessageText, "연금/저축 소득공제");
                                mAppInterface.OnAppMessageEvent(vMessageText);
                                System.Windows.Forms.Application.DoEvents();

                                vPrintingLine = 15;
                                for (int vRow3 = 0; vRow3 < pGR_PRINT_SAVING_INFO_2.RowCount; vRow3++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_2, vRow3, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                vPrintingLine = 26; //출력되지 않은 수 만큼 더함, 다음 출력 위치를 위해
                                for (int vRow4 = 0; vRow4 < pGR_PRINT_SAVING_INFO_3.RowCount; vRow4++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_3, vRow4, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                vPrintingLine = 37;
                                for (int vRow5 = 0; vRow5 < pGR_PRINT_SAVING_INFO_4.RowCount; vRow5++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_4, vRow5, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                vPrintingLine = 48;
                                for (int vRow6 = 0; vRow6 < pGR_PRINT_SAVING_INFO_5.RowCount; vRow6++)
                                {
                                    vPrintingLine = XLLine4(pGR_PRINT_SAVING_INFO_5, vRow6, vPrintingLine, vGDColumn_4, vXLColumn_4);
                                }
                                mCopyLineSUM = CopyAndPaste_2(mCopyLineSUM);
                            }
                            else if (pPrint_Year == "2013")
                            {
                                vPrintingLine = 4;
                                vPrintingLine = XLHeader_PRINT_SAVING_INFO_2013(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_5);

                                vMessageText = string.Format("{0} - {1}", vMessageText, "연금/저축 소득공제");
                                mAppInterface.OnAppMessageEvent(vMessageText);
                                System.Windows.Forms.Application.DoEvents();


                                vPrintingLine = 16;
                                for (int vRow3 = 0; vRow3 < pGR_PRINT_SAVING_INFO_2.RowCount; vRow3++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_2, vRow3, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                vPrintingLine = 29; //출력되지 않은 수 만큼 더함, 다음 출력 위치를 위해
                                for (int vRow4 = 0; vRow4 < pGR_PRINT_SAVING_INFO_3.RowCount; vRow4++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_3, vRow4, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                vPrintingLine = 43;
                                for (int vRow5 = 0; vRow5 < pGR_PRINT_SAVING_INFO_4.RowCount; vRow5++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_4, vRow5, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                mCopyLineSUM = CopyAndPaste_2(mCopyLineSUM);
                            }
                            else
                            {
                                vPrintingLine = 4;
                                vPrintingLine = XLHeader_PRINT_SAVING_INFO_2013(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_5);

                                vMessageText = string.Format("{0} - {1}", vMessageText, "연금/저축 소득공제");
                                mAppInterface.OnAppMessageEvent(vMessageText);
                                System.Windows.Forms.Application.DoEvents();


                                vPrintingLine = 16;
                                for (int vRow3 = 0; vRow3 < pGR_PRINT_SAVING_INFO_2.RowCount; vRow3++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_2, vRow3, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                vPrintingLine = 26; //출력되지 않은 수 만큼 더함, 다음 출력 위치를 위해
                                for (int vRow4 = 0; vRow4 < pGR_PRINT_SAVING_INFO_3.RowCount; vRow4++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_3, vRow4, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                vPrintingLine = 46;
                                for (int vRow5 = 0; vRow5 < pGR_PRINT_SAVING_INFO_4.RowCount; vRow5++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_4, vRow5, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                vPrintingLine = 36;
                                for (int vRow5 = 0; vRow5 < pGR_PRINT_SAVING_INFO_4.RowCount; vRow5++)
                                {
                                    vPrintingLine = XLLine3(pGR_PRINT_SAVING_INFO_6, vRow5, vPrintingLine, vGDColumn_3, vXLColumn_3);
                                }

                                mCopyLineSUM = CopyAndPaste_2(mCopyLineSUM);
                            }

                            //월세액/거주지간 소득공제
                            if (pPRINT_HOUSE_YN == "Y")
                            {
                                //연금/저축 등 소득공재 명세서 출력할 것이 있고
                                //월세액/거주지간 소득공제 명세서 출력할 것이 있을 경우
                                if (pidaHOUSE_LEASE_INFO_10.CurrentRows.Count > 0 || pidaHOUSE_LEASE_INFO_20.CurrentRows.Count > 0)
                                {
                                    if (pPrint_Year == "2011" || pPrint_Year == "2012" || pPrint_Year == "2013")
                                    {
                                        //둘다 있을경우 PAGE 수는 5로 증가 
                                        pPrintPage = 5;
                                        vPrintingLine = 4;
                                        vPrintingLine = XLHeader_PRINT_HOUSE_INFO_2013(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_5);

                                        vMessageText = string.Format("{0} - {1}", vMessageText, "월세액/거주지간 소득공제");
                                        mAppInterface.OnAppMessageEvent(vMessageText);
                                        System.Windows.Forms.Application.DoEvents();

                                        vPrintingLine = 14;

                                        vPrintingLine = XLLine5(pidaHOUSE_LEASE_INFO_10, vPrintingLine);

                                        vPrintingLine = 29;

                                        vPrintingLine = XLLine6(pidaHOUSE_LEASE_INFO_20, vPrintingLine);

                                        mCopyLineSUM = CopyAndPaste_3(mCopyLineSUM);
                                    }
                                    else //2014
                                    {
                                        //둘다 있을경우 PAGE 수는 5로 증가 
                                        pPrintPage = 5;
                                        vPrintingLine = 4;
                                        vPrintingLine = XLHeader_PRINT_HOUSE_INFO_2013(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_5);

                                        vMessageText = string.Format("{0} - {1}", vMessageText, "월세액/거주지간 소득공제");
                                        mAppInterface.OnAppMessageEvent(vMessageText);
                                        System.Windows.Forms.Application.DoEvents();

                                        vPrintingLine = 14;

                                        vPrintingLine = XLLine7(pidaHOUSE_LEASE_INFO_10, vPrintingLine);

                                        vPrintingLine = 29;

                                        vPrintingLine = XLLine8(pidaHOUSE_LEASE_INFO_20, vPrintingLine);

                                        mCopyLineSUM = CopyAndPaste_3(mCopyLineSUM);
                                    }
                                }
                            }
                        }
                        else
                        {
                            pPrintPage = 3;
                            //월세액/거주지간 소득공제 명세서만 있을경우
                            if (pPRINT_HOUSE_YN == "Y")
                            {
                                if (pidaHOUSE_LEASE_INFO_10.CurrentRows.Count > 0 || pidaHOUSE_LEASE_INFO_20.CurrentRows.Count > 0)
                                {
                                    mCopyLineSUM++;

                                    if (pPrint_Year == "2011" || pPrint_Year == "2012" || pPrint_Year == "2013")
                                    {
                                        pPrintPage = 4;
                                        vPrintingLine = 4;
                                        vPrintingLine = XLHeader_PRINT_HOUSE_INFO_2013(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_5);

                                        vMessageText = string.Format("{0} - {1}", vMessageText, "월세액/거주지간 소득공제");
                                        mAppInterface.OnAppMessageEvent(vMessageText);
                                        System.Windows.Forms.Application.DoEvents();

                                        vPrintingLine = 14;

                                        vPrintingLine = XLLine5(pidaHOUSE_LEASE_INFO_10, vPrintingLine);

                                        vPrintingLine = 29;

                                        vPrintingLine = XLLine6(pidaHOUSE_LEASE_INFO_20, vPrintingLine);

                                        mCopyLineSUM = CopyAndPaste_3(mCopyLineSUM);
                                    }
                                    else //2014
                                    {
                                        pPrintPage = 4;
                                        vPrintingLine = 4;
                                        vPrintingLine = XLHeader_PRINT_HOUSE_INFO_2013(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_5);

                                        vMessageText = string.Format("{0} - {1}", vMessageText, "월세액/거주지간 소득공제");
                                        mAppInterface.OnAppMessageEvent(vMessageText);
                                        System.Windows.Forms.Application.DoEvents();

                                        vPrintingLine = 14;

                                        vPrintingLine = XLLine7(pidaHOUSE_LEASE_INFO_10, vPrintingLine);

                                        vPrintingLine = 29;

                                        vPrintingLine = XLLine8(pidaHOUSE_LEASE_INFO_20, vPrintingLine);

                                        mCopyLineSUM = CopyAndPaste_3(mCopyLineSUM);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        pPrintPage = 3;
                        //월세액/거주지간 소득공제 명세서만 있을경우
                        if (pPRINT_HOUSE_YN == "Y")
                        {
                            if (pidaHOUSE_LEASE_INFO_10.CurrentRows.Count > 0 || pidaHOUSE_LEASE_INFO_20.CurrentRows.Count > 0)
                            {
                                mCopyLineSUM++;

                                pPrintPage = 4;
                                vPrintingLine = 4;
                                vPrintingLine = XLHeader_PRINT_HOUSE_INFO_2013(pGrid_WITHHOLDING_TAX_13, vRow1, vPrintingLine, vGDColumn_5);

                                vMessageText = string.Format("{0} - {1}", vMessageText, "월세액/거주지간 소득공제");
                                mAppInterface.OnAppMessageEvent(vMessageText);
                                System.Windows.Forms.Application.DoEvents();

                                vPrintingLine = 14;

                                vPrintingLine = XLLine5(pidaHOUSE_LEASE_INFO_10, vPrintingLine);

                                vPrintingLine = 29;

                                vPrintingLine = XLLine6(pidaHOUSE_LEASE_INFO_20, vPrintingLine);

                                mCopyLineSUM = CopyAndPaste_3(mCopyLineSUM);
                            }
                        }
                    }

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////
                    if (pOutChoice == "PRINT")
                    {
                        if (pPrintPage == 0)
                        {
                            Printing(1, mPageNumber);
                        }
                        else if (pPrintPage == 1)
                        {
                            Printing(1, 1);
                        }
                        else if (pPrintPage == 2)
                        {
                            Printing(1, 2);
                        }
                        else if (pPrintPage == 3)
                        {
                            Printing(1, 3);
                        }
                        else if (pPrintPage == 4)
                        {
                            Printing(1, 4);
                        }
                        else if (pPrintPage == 5)
                        {
                            Printing(1, 5);
                        }
                        //else if (pPrintPage == 6)
                        //{
                        //    Printing(1, 6);
                        //}
                        //else if (pPrintPage == 7)
                        //{
                        //    Printing(1, 7);
                        //}
                        //else if (pPrintPage == 8)
                        //{
                        //    Printing(1, 8);
                        //}

                        vMessageText = string.Format("{0} - {1}", vMessageText, pPrintPage);
                        mAppInterface.OnAppMessageEvent(vMessageText);
                        System.Windows.Forms.Application.DoEvents();
                    }

                    if (pOutChoice == "PDF")
                    {

                        DeleteSheet();
                        PDF(vSaveFileName);  //PDF 파일명
                    }
                }
                mPrinting.XLOpenFileClose();

            }

            return mPageNumber;
        }
         
        #endregion;

        #region ----- Copy&Paste Sheet Method 1 ----

        //첫번째 페이지 복사
        private int CopyAndPaste_1(int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX_1;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            int vCopyPrintingRowSTART = 1;
            int vCopyPrintingRowEnd = mIncrementCopyMAX_1;

            mPrinting.XLActiveSheet("SourceTab1");

            object vRangeSource = mPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX_1, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            mPrinting.XLActiveSheet("Destination");
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber = mPageNumber + 3; //페이지 번호

            return vCopyPrintingRowEnd;
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method 2 ----

        //두번째 페이지 복사
        private int CopyAndPaste_2(int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX_2;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            int vCopyPrintingRowEnd = vCopyPrintingRowSTART + mIncrementCopyMAX_2;

            mPrinting.XLActiveSheet("SourceTab2");
            object vRangeSource = mPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX_2, mCopyColumnEND);

            mPrinting.XLActiveSheet("Destination");
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND);
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber++; //페이지 번호
            return vCopyPrintingRowEnd;
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method 3 ----

        //두번째 페이지 복사
        private int CopyAndPaste_3(int pCopySumPrintingLine)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX_2;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            int vCopyPrintingRowEnd = vCopyPrintingRowSTART + mIncrementCopyMAX_3;

            mPrinting.XLActiveSheet("SourceTab3");
            object vRangeSource = mPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX_3, mCopyColumnEND);

            mPrinting.XLActiveSheet("Destination");
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND);
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber++; //페이지 번호

            return vCopyPrintingRowEnd;
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

            vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder.ToString(), vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;

        #region ----- PDF Method ----

        //public void PDF(string pSaveFileName)
        //{
        //    System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

        //    int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
        //    vMaxNumber = vMaxNumber + 1;
        //    string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

        //    vSaveFileName = string.Format("{0}\\{1}.pdf", vWallpaperFolder, vSaveFileName);
        //    bool isSuccess = mPrinting.XLSaveAs_PDF(vSaveFileName);
        //    string vMessage = mPrinting.MessageError;
        //    int tmp = vMaxNumber;
        //}

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
                isSuccess = mPrinting.XLDeleteSheet("SourceTab1");
                isSuccess = mPrinting.XLDeleteSheet("SourceTab2");
                isSuccess = mPrinting.XLDeleteSheet("SourceTab3");
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