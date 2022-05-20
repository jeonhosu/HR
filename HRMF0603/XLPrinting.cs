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

namespace HRMF0603
{
    public class XLPrinting1
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mPageNumber = 0;

        private string mXLOpenFileName = string.Empty;

        private int mPrintingLineSTART = 1;  //Line

        private int mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치, 복사 행 누적
        private int mIncrementCopyMAX = 62; //복사되어질 행의 범위

        private int mCopyColumnSTART = 1;    //복사되어  진 행 누적 수
        private int mCopyColumnEND = 43;     //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

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

        public XLPrinting1(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
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

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[82];
            pXLColumn = new int[82];

            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[0] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DEPT_NAME");                // 부서                                                 
            pGDColumn[1] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PERSON_NUM");               // 사번                                                 
            pGDColumn[2] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("NAME");                     // 성명                                                 
            pGDColumn[3] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REPRE_NUM");                // 주민번호                                             
            pGDColumn[4] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("JOIN_DATE");                // 입사일자                                             
            pGDColumn[5] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("EXPIRE_DATE");              // 최종중간정산일                                       
            pGDColumn[6] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_DATE");              // 퇴직일자                                             
            pGDColumn[7] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("CONTINUE_DAY");             // 근무일수                                             
            pGDColumn[8] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ADDRESS");                  // 주소                                                 
            pGDColumn[9] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_DAY");                 // 산정일수                                             
            pGDColumn[10] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DED_DAY");                 // 제외일수                                             
            pGDColumn[11] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_NAME");             // 퇴직사유                                             
            pGDColumn[12] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ADJUSTMENT_TYPE");         // 정산구분(R : 퇴직금, N : 중도정산)                   
            pGDColumn[13] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE4");             // 시작 날짜1                                           
            pGDColumn[14] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE3");             // 시작 날짜2                                           
            pGDColumn[15] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE2");             // 시작 날짜3                                           
            pGDColumn[16] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE1");             // 시작 날짜4                                           
            pGDColumn[17] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE4");               // 마지막 날짜1                                         
            pGDColumn[18] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE3");               // 마지막 날짜2                                         
            pGDColumn[19] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE2");               // 마지막 날짜3                                         
            pGDColumn[20] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE1");               // 마지막 날짜4                                         
            pGDColumn[21] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_AMOUNT4");           // 급여1                                                
            pGDColumn[22] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_AMOUNT3");           // 급여2                                                
            pGDColumn[23] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_AMOUNT2");           // 급여3                                                
            pGDColumn[24] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_AMOUNT1");           // 급여4                                                
            pGDColumn[25] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_PAY_AMOUNT");        // 급여(합계)                                           
            pGDColumn[26] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY4");                    // 근무일수1                                            
            pGDColumn[27] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY3");                    // 근무일수2                                            
            pGDColumn[28] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY2");                    // 근무일수3                                            
            pGDColumn[29] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY1");                    // 근무일수4                                            
            pGDColumn[30] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY_3RD_COUNT");           // 근무일수(합계)                                       
            pGDColumn[31] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_BOUNS");             // 연간 상여 계                                         
            pGDColumn[32] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_BONUS_AMOUNT");      // 상여(합계)                                           
            pGDColumn[33] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("YEAR_ALLOWANCE");          // 연간 연월차 계                                       
            pGDColumn[34] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("YEAR_ALLOWANCE_AMOUNT");   // 연월차(합계)                                         
            pGDColumn[35] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_PAYMENT");           // 임금 총액                                            
            pGDColumn[36] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY_AVR_PAY");             // 일 평균 임금                                         
            pGDColumn[37] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_AMOUNT");           // 퇴직급여                                             
            pGDColumn[38] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_HONORARY_AMOUNT_1"); // 명예퇴직수당 등                                      
            pGDColumn[39] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_INSUR_AMOUNT");     // 퇴직보험금 등                                        
            pGDColumn[40] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PRE_CORP_NAME");           // 지급처명-종전근무지내역                              
            pGDColumn[41] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PRE_RETIRE_AMOUNT");       // 퇴직급여액-종전근무지내역                            
            pGDColumn[42] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PRE_INCOME_AMOUNT");       // 소득세-종전근무지내역                                
            pGDColumn[43] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PRE_LOCAL_AMOUNT");        // 주민세-종전근무지내역                                
            pGDColumn[44] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_PERIOD");             // 근속기간-종전근무지내역                              
            pGDColumn[45] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_MONTH");              // 근속월수-종전근무지내역                              
            pGDColumn[46] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DUPL_MONTH");              // 중복월수-종전근무지내역                              
            pGDColumn[47] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_TOTAL_AMOUNT");     // 퇴직급여액(법정 퇴직 급여)                           
            pGDColumn[48] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_HONORARY_AMOUNT_2"); // 퇴직급여액(법정이외 퇴직급여)                        
            pGDColumn[49] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_DED_RATE1");        // 퇴직급여공제(법정퇴직급여 및 법정이외 퇴직급여)      
            pGDColumn[50] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_DED_AMOUNT");       // 퇴직소득공제 - 퇴직급여공제 - 법정퇴직급여           
            pGDColumn[51] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_DED_RATE2");        // 퇴직급여공제(법정퇴직급여 및 법정이외 퇴직급여)      
            pGDColumn[52] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_INCOME_DED_AMOUNT");     // 퇴직소득공제 - 퇴직급여공제 - 법정이외 퇴직급여      
            pGDColumn[53] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_YEAR");               // 근속연수공제(법정퇴직급여)                           
            pGDColumn[54] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_DED_AMOUNT");         // 퇴직소득공제 - 근속연수공제 - 법정퇴직급여           
            pGDColumn[55] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_LONG_YEAR");             // 근속연수공제(법정이외 퇴직급여)                      
            pGDColumn[56] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_LONG_DED_AMOUNT");       // 퇴직소득공제 - 근속연수공제 - 법정이외 퇴직급여      
            pGDColumn[57] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DED_SUM_AMOUNT");          // 퇴직소득공제 - 계 - 법정퇴직급여                     
            pGDColumn[58] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_DED_SUM_AMOUNT");        // 퇴직소득공제 - 계 - 법정이외 퇴직급여                
            pGDColumn[59] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TAX_STD_AMOUNT");          // 세액계산근거 - 퇴직소득과세표준 - 법정퇴직급여       
            pGDColumn[60] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_TAX_STD_AMOUNT");        // 세액계산근거 - 퇴직소득과세표준 - 법정이외 퇴직급여  
            pGDColumn[61] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("AVG_TAX_STD_AMOUNT");      // 세액계산근거 - 연평균과세표준 - 법정퇴직급여         
            pGDColumn[62] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_AVG_TAX_STD_AMOUNT");    // 세액계산근거 - 연평균과세표준 - 법정이외 퇴직급여    
            pGDColumn[63] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TAX_RATE");                // 퇴직소득세율                                         
            pGDColumn[64] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("AVG_COMP_TAX_AMOUNT");     // 세액계산근거 - 연평균산출세액 - 법정퇴직급여         
            pGDColumn[65] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_AVG_COMP_TAX_AMOUNT");   // 세액계산근거 - 연평균산출세액 - 법정이외 퇴직급여    
            pGDColumn[66] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("COMP_TAX_AMOUNT");         // 세액계산근거 - 산출세액 - 법정퇴직급여               
            pGDColumn[67] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_COMP_TAX_AMOUNT");       // 세액계산근거 - 산출세액 - 법정이외 퇴직급여          
            pGDColumn[68] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TAX_DED_AMOUNT");          // 세액계산근거 - 세액공제(외국납부) - 법정퇴직급여     
            pGDColumn[69] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_TAX_DED_AMOUNT");        // 세액계산근거 - 세액공제(외국납부) - 법정이외 퇴직급여
            pGDColumn[70] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_TAX_AMOUNT");       // 결정세액 - 퇴직소득세 - 법정퇴직급여                 
            pGDColumn[71] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_INCOME_TAX_AMOUNT");     // 결정세액 - 퇴직소득세 - 법정이외 퇴직급여            
            pGDColumn[72] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RESIDENT_TAX_AMOUNT");     // 결정세액 - 퇴직주민세                                
            pGDColumn[73] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REAL_TOTAL_AMOUNT");       // 세후 지급액                                          
            pGDColumn[74] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ETC_SUPP_AMOUNT");         // 공제 및 가산항목(기타공제 사유)                      
            pGDColumn[75] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ETC_DED_AMOUNT");          // 공제 및 가산항목(기타공제)                           
            pGDColumn[76] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_TOTAL_AMOUNT1");    // 차인직급액                                           
            pGDColumn[77] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("BANK_NAME");               // 은행                                                 
            pGDColumn[78] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("BANK_ACCOUNTS");           // 계좌번호                                             
            pGDColumn[79] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("NAME1");                   // 영수인
           
            //---------------------------------------------------------------------------------------------------------------------
            pXLColumn[0]  =  6; // 부서                                                 
            pXLColumn[1]  = 16; // 사번                                                 
            pXLColumn[2]  = 26; // 성명                                                 
            pXLColumn[3]  = 36; // 주민번호                                             
            pXLColumn[4]  =  6; // 입사일자                                             
            pXLColumn[5]  = 16; // 최종중간정산일                                       
            pXLColumn[6]  = 26; // 퇴직일자                                             
            pXLColumn[7]  = 36; // 근무일수                                             
            pXLColumn[8]  =  6; // 주소                                                 
            pXLColumn[9]  = 26; // 산정일수                                             
            pXLColumn[10] = 36; // 제외일수                                             
            pXLColumn[11] = 26; // 퇴직사유                                             
            pXLColumn[12] = 38; // 정산구분(R : 퇴직금, N : 중도정산)                   
            pXLColumn[13] = 16; // 시작 날짜1                                           
            pXLColumn[14] = 21; // 시작 날짜2                                           
            pXLColumn[15] = 26; // 시작 날짜3                                           
            pXLColumn[16] = 31; // 시작 날짜4                                           
            pXLColumn[17] = 16; // 마지막 날짜1                                         
            pXLColumn[18] = 21; // 마지막 날짜2                                         
            pXLColumn[19] = 26; // 마지막 날짜3                                         
            pXLColumn[20] = 31; // 마지막 날짜4                                         
            pXLColumn[21] = 16; // 급여1                                                
            pXLColumn[22] = 21; // 급여2                                                
            pXLColumn[23] = 26; // 급여3                                                
            pXLColumn[24] = 31; // 급여4                                                
            pXLColumn[25] = 36; // 급여(합계)                                           
            pXLColumn[26] = 16; // 근무일수1                                            
            pXLColumn[27] = 21; // 근무일수2                                            
            pXLColumn[28] = 26; // 근무일수3                                            
            pXLColumn[29] = 31; // 근무일수4                                            
            pXLColumn[30] = 36; // 근무일수(합계)                                       
            pXLColumn[31] = 16; // 연간 상여 계                                         
            pXLColumn[32] = 36; // 상여(합계)                                           
            pXLColumn[33] = 16; // 연간 연월차 계                                       
            pXLColumn[34] = 36; // 연월차(합계)                                         
            pXLColumn[35] = 36; // 임금 총액                                            
            pXLColumn[36] = 36; // 일 평균 임금                                         
            pXLColumn[37] = 36; // 퇴직급여                                             
            pXLColumn[38] = 36; // 명예퇴직수당 등                                      
            pXLColumn[39] = 36; // 퇴직보험금 등                                        
            pXLColumn[40] =  1; // 지급처명-종전근무지내역                              
            pXLColumn[41] = 10; // 퇴직급여액-종전근무지내역                            
            pXLColumn[42] = 16; // 소득세-종전근무지내역                                
            pXLColumn[43] = 21; // 주민세-종전근무지내역                                
            pXLColumn[44] = 26; // 근속기간-종전근무지내역                              
            pXLColumn[45] = 36; // 근속월수-종전근무지내역                              
            pXLColumn[46] = 40; // 중복월수-종전근무지내역                              
            pXLColumn[47] = 28; // 퇴직급여액(법정 퇴직 급여)                           
            pXLColumn[48] = 36; // 퇴직급여액(법정이외 퇴직급여)                        
            pXLColumn[49] = 28; // 퇴직급여공제(법정퇴직급여 및 법정이외 퇴직급여)      
            pXLColumn[50] = 31; // 퇴직소득공제 - 퇴직급여공제 - 법정퇴직급여           
            pXLColumn[51] = 36; // 퇴직급여공제(법정퇴직급여 및 법정이외 퇴직급여)      
            pXLColumn[52] = 39; // 퇴직소득공제 - 퇴직급여공제 - 법정이외 퇴직급여      
            pXLColumn[53] = 28; // 근속연수공제(법정퇴직급여)                           
            pXLColumn[54] = 31; // 퇴직소득공제 - 근속연수공제 - 법정퇴직급여           
            pXLColumn[55] = 36; // 근속연수공제(법정이외 퇴직급여)                      
            pXLColumn[56] = 39; // 퇴직소득공제 - 근속연수공제 - 법정이외 퇴직급여      
            pXLColumn[57] = 28; // 퇴직소득공제 - 계 - 법정퇴직급여                     
            pXLColumn[58] = 36; // 퇴직소득공제 - 계 - 법정이외 퇴직급여                
            pXLColumn[59] = 28; // 세액계산근거 - 퇴직소득과세표준 - 법정퇴직급여       
            pXLColumn[60] = 36; // 세액계산근거 - 퇴직소득과세표준 - 법정이외 퇴직급여  
            pXLColumn[61] = 28; // 세액계산근거 - 연평균과세표준 - 법정퇴직급여         
            pXLColumn[62] = 36; // 세액계산근거 - 연평균과세표준 - 법정이외 퇴직급여    
            pXLColumn[63] = 28; // 퇴직소득세율                                         
            pXLColumn[64] = 28; // 세액계산근거 - 연평균산출세액 - 법정퇴직급여         
            pXLColumn[65] = 36; // 세액계산근거 - 연평균산출세액 - 법정이외 퇴직급여    
            pXLColumn[66] = 28; // 세액계산근거 - 산출세액 - 법정퇴직급여               
            pXLColumn[67] = 36; // 세액계산근거 - 산출세액 - 법정이외 퇴직급여          
            pXLColumn[68] = 28; // 세액계산근거 - 세액공제(외국납부) - 법정퇴직급여     
            pXLColumn[69] = 36; // 세액계산근거 - 세액공제(외국납부) - 법정이외 퇴직급여
            pXLColumn[70] = 28; // 결정세액 - 퇴직소득세 - 법정퇴직급여                 
            pXLColumn[71] = 36; // 결정세액 - 퇴직소득세 - 법정이외 퇴직급여            
            pXLColumn[72] = 28; // 결정세액 - 퇴직주민세                                
            pXLColumn[73] = 28; // 세후 지급액                                          
            pXLColumn[74] = 16; // 공제 및 가산항목(기타공제 사유)                      
            pXLColumn[75] = 36; // 공제 및 가산항목(기타공제)                           
            pXLColumn[76] = 36; // 차인지급액                                           
            pXLColumn[77] =  6; // 은행                                                 
            pXLColumn[78] = 21; // 계좌번호                                             
            pXLColumn[79] = 33; // 영수인
        }

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_ALLOWANCE, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[90];
            pXLColumn = new int[90];

            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[0] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PERSON_NUM");                  // 사번                                              
            pGDColumn[1] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("OPERATING_UNIT_NAME");    //사업장                 
            pGDColumn[2] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("JOIN_DATE");                      // 입사일                            
            pGDColumn[3] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("CONTINUE_DAY");                    // 근속기간      
                                                   
            pGDColumn[4] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("NAME");                // 성명                                             
            pGDColumn[5] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DEPT_NAME");              // 부서
            pGDColumn[6] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_DATE_FR");              // 정산시작일(기산일)
            pGDColumn[7] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DED_DAY");          // 휴직기간

            pGDColumn[8] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REPRE_NUM");                  //주민번호
            pGDColumn[9] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ADDRESS");                 // 주소                                             
            pGDColumn[10] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_DATE");                 // 퇴사일                                             
            pGDColumn[11] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("BANK_ACCOUNTS");             // 계좌번호

            pGDColumn[12] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("POST_NAME");         // 직위                 
            pGDColumn[40] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("CLOSED_DATE");         // 지급일                 

            pGDColumn[13] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE4");             // 시작 날짜1                                           
            pGDColumn[14] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE3");             // 시작 날짜2                                           
            pGDColumn[15] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE2");             // 시작 날짜3                                           
            pGDColumn[16] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE1");             // 시작 날짜4                                           
            pGDColumn[17] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE4");               // 마지막 날짜1                                         
            pGDColumn[18] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE3");               // 마지막 날짜2                                         
            pGDColumn[19] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE2");               // 마지막 날짜3                                         
            pGDColumn[20] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE1");               // 마지막 날짜4                                             
            pGDColumn[26] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY4");                    // 근무일수1                               
            pGDColumn[27] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY3");                    // 근무일수2                               
            pGDColumn[28] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY2");                    // 근무일수3                               
            pGDColumn[29] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY1");                    // 근무일수4                               
            pGDColumn[30] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY_3RD_COUNT");           // 근무일수(합계)       
                                            
            pGDColumn[33] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_PAY_AMOUNT");        // 산정급여 
            pGDColumn[34] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("YEAR_BONUS_AMOUNT");     //산정상여/연차
            pGDColumn[35] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_PAYMENT");          //합계
            pGDColumn[36] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY_AVR_PAY");                 // 일 평균 임금                                         
            pGDColumn[37] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_AMOUNT");             // 퇴직급여                    
                                     
            pGDColumn[38] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_AMOUNT");            // 퇴직금                                   
            pGDColumn[39] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("GLORY_AMOUNT");            // 위로금                            
            pGDColumn[47] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_GROUP_AMOUNT");                  // 단체퇴직보험금  ?????
            pGDColumn[48] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("HONORARY_AMOUNT");       // 명예퇴직수당등                     
            pGDColumn[49] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_PENSION_AMOUNT");                  //퇴직연금 ???
            pGDColumn[50] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_TAX_FREE_AMOUNT");                  //비과세소득  ????        
            pGDColumn[51] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_AMOUNT");                  //지급총액 ????

            pGDColumn[52] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_TAX_AMOUNT_3");              //퇴직소득세     
            pGDColumn[53] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RESIDENT_TAX_AMOUNT");             //퇴직지방소득세
            pGDColumn[54] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ETC_DED_AMOUNT");                     //기타공제 
            pGDColumn[55] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_DED_AMOUNT");           // 공제총액 
            pGDColumn[56] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_TOTAL_AMOUNT1");            //차인지급액 
                                         
            pGDColumn[77] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("BANK_NAME");            // 은행                                                   
            pGDColumn[79] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("NAME1");                   // 영수인

            ///////////////////////////LINE 부분
            pGDColumn[80] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("RETIRE_SALARY_ITEM_NAME");  //항목명                 
            pGDColumn[81] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT1");          //마지막달 -3     
            pGDColumn[82] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT2");          //마지막달 -2
            pGDColumn[83] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT3");          //마지막달 -1        
            pGDColumn[84] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT4");          //마지막달 
            pGDColumn[85] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT");            //총금액

            pGDColumn[86] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("SUB_INCOME_TAX_AMOUNT");              //퇴직소득세     
            pGDColumn[87] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("SUB_RESIDENT_TAX_AMOUNT");             //퇴직지방소득세
            pGDColumn[88] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REAL_TOTAL_DED_AMOUNT");             //퇴직지방소득세
            pGDColumn[89] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REAL_RETIRE_TOTAL_AMOUNT");            //차인지급액 


            //---------------------------------------------------------------------------------------------------------------------
            /////퇴직자 인적사항
            pXLColumn[0] = 6; // 사번                                               
            pXLColumn[1] = 6; // 사업장                                        
            pXLColumn[2] = 6; // 입사일                                          
            pXLColumn[3] = 6; // 근속기간 
                                                 
            pXLColumn[4] = 18; // 성명                                           
            pXLColumn[5] = 18; //부서
            pXLColumn[6] = 18; // 기산일
            pXLColumn[7] = 18; //휴직기간 

            pXLColumn[8] = 28; // 주민번호
            pXLColumn[9] = 28; // 주소                                          
            pXLColumn[10] = 28; // 퇴사일                                        
            pXLColumn[11] = 33; // 계좌번호                         
                         
            pXLColumn[12] = 38; // 직위    
          //  pXLColumn[12] = 38; // 지급일 

            ///// 퇴직금 산정내역 헤더
            pXLColumn[16] = 8; // 시작 날짜1
            pXLColumn[15] = 15; // 시작 날짜2                                           
            pXLColumn[14] = 22; // 시작 날짜3                                           
            pXLColumn[13] = 29; // 시작 날짜4              
                                         
            pXLColumn[20] = 8; // 마지막 날짜1                                         
            pXLColumn[19] = 15;// 마지막 날짜2                                         
            pXLColumn[18] = 22; // 마지막 날짜3                                         
            pXLColumn[17] = 29; // 마지막 날짜4           

            pXLColumn[29] = 8; // 근무일수1                                            
            pXLColumn[28] = 15; // 근무일수2                                            
            pXLColumn[27] = 22; // 근무일수3                                            
            pXLColumn[26] = 29; // 근무일수4                                            
            pXLColumn[30] = 36; // 근무일수(합계) 
                                                                        
            pXLColumn[33] = 1; // 산정급여                               
            pXLColumn[34] = 10; // 산정상여+연차
            pXLColumn[35] = 18; // 합계                                     
            pXLColumn[36] = 28; // 일 평균 임금                                         
            pXLColumn[37] = 36; // 퇴직급여           
                                              
            pXLColumn[38] = 13;  //퇴직금지급내역 금액    
            pXLColumn[39] = 35;  //퇴직금공제내역 금액

            pXLColumn[77] = 28;  //은행명 

            //////LINE부분
            pXLColumn[80] = 1; // 급여항목
            pXLColumn[81] = 8; // 마지막달 -3                                   
            pXLColumn[82] = 15; // 마지막달 -2                                
            pXLColumn[83] = 22; // 마지막달 -1                                        
            pXLColumn[84] = 29; // 마지막달                                         
            pXLColumn[85] = 36; // 합계 
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

        #region ----- Excel Wirte [Header] Methods ----
        private int HeaderWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, int pGridRow, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = 8; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;


            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                
                // 사번 
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 사업장 
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }

                // 입사일
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, DBNull.Value);
                }

                // 근속기간 
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine+6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+6, vXLColumnIndex, vConvertString);
                }
                

                // 성명
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    mPrinting.XLSetCell(80, 28, vConvertString);
                    mPrinting.XLSetCell(78, 26, DateTime.Today );
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 부서
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }

                // 기산일
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, DBNull.Value);
                }

                // 휴직기간 
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                
                // 주민번호
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 주소
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }

                // 퇴사일
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, DBNull.Value);
                }

                // 은행
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", "("+vConvertString+")");
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                // 계좌번호
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }

                // 직위
                vGDColumnIndex = pGDColumn[12];  
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {                 
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine , vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine , vXLColumnIndex, vConvertString);
                }

                // 지급일
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, DBNull.Value);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 13;
                //-------------------------------------------------------------------

                // 산정급여 
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 산정상여/연차
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 합계 
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 일평균  
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 퇴직급여 
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 시작 날짜1 
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 시작 날짜2 
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 시작 날짜3 
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 시작 날짜4
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 마지막 날짜1
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 마지막 날짜2
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 마지막 날짜3
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 마지막 날짜4
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 근무일수1 
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 근무일수2
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무일수3
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무일수4
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무일수합계
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 31;
                //-------------------------------------------------------------------

                // 퇴직금  
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 퇴직소득세  
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 위로금 
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 퇴직지방소득 
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 단체퇴직보험금  
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 기타공제
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 명예퇴직수당 
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                // 퇴직연금 
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine+4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+4, vXLColumnIndex, vConvertString);
                }
                // 비과세소득 
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine+6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+6, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 8;
                //-------------------------------------------------------------------
                // 지급총액 
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // 공제총액  
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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
                // 차인지급액  
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
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
            

            return vXLine;
        }

        #endregion

        #region ----- Line Write Method 1 -----
        private int XLLine1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_ETC_ALLOWANCE, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pCourse)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            int vGDColumnIndex1 = 0;
            int vXLColumnIndex1 = 0;
            int vGDColumnIndex2 = 0;
            int vXLColumnIndex2 = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                int vCountRow = pGrid_ETC_ALLOWANCE.RowCount; //pGrid_ETC_ALLOWANCE 그리드의 총 행수

                mPrinting.XLActiveSheet("Destination");

                //-------------------------------------------------------------------
                vXLine = vXLine + 5;
                //-------------------------------------------------------------------

                // 부서
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 사번
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 주민등록번호
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 입사일자
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 최종중간정산일
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 퇴직일자
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 근무일수
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 산정일수
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외일수
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
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

                // 퇴직사유
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 정산구분
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "R")
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "●");
                    }
                    else
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 4), "●");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 4), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // 시작 날짜1 
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 시작 날짜2 
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 시작 날짜3 
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 시작 날짜4
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 마지막 날짜1
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 마지막 날짜2
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 마지막 날짜3
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // 마지막 날짜4
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 급여1
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 급여2
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 급여3
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 급여4
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 급여(합계) 
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 근무일수1
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무일수2
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무일수3
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무일수4
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무일수(합계)
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,##0}", vConvertDecimal);
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

                // 연간 상여 계
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 상여(합계)
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 연간 연월차 계
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 연월차(합계) 
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 임금 총액
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 일 평균 임금  
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직급여
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 명예퇴직수당 등
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직보험금 등 
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 지급처명-종전근무지내역
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여액-종전근무지내역
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 소득세-종전근무지내역
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 주민세-종전근무지내역
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속기간-종전근무지내역 
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수-종전근무지내역
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 중복월수-종전근무지내역  
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 퇴직급여액(법정 퇴직 급여)    
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여액(법정이외 퇴직급여) 
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직급여공제(법정퇴직급여 및 법정이외 퇴직급여) 
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득공제 - 퇴직급여공제 - 법정퇴직급여 
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여공제(법정퇴직급여 및 법정이외 퇴직급여)
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득공제 - 퇴직급여공제 - 법정이외 퇴직급여
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 근속연수공제(법정퇴직급여)
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득공제 - 근속연수공제 - 법정퇴직급여
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속연수공제(법정이외 퇴직급여)
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득공제 - 근속연수공제 - 법정이외 퇴직급여
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직소득공제 - 계 - 법정퇴직급여 
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득공제 - 계 - 법정이외 퇴직급여
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 세액계산근거 - 퇴직소득과세표준 - 법정퇴직급여
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 세액계산근거 - 퇴직소득과세표준 - 법정이외 퇴직급여
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 세액계산근거 - 연평균과세표준 - 법정퇴직급여
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 세액계산근거 - 연평균과세표준 - 법정이외 퇴직급여
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직소득세율
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 세액계산근거 - 연평균산출세액 - 법정퇴직급여
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 세액계산근거 - 연평균산출세액 - 법정이외 퇴직급여
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 세액계산근거 - 산출세액 - 법정퇴직급여    
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 세액계산근거 - 산출세액 - 법정이외 퇴직급여
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                //  세액계산근거 - 세액공제(외국납부) - 법정퇴직급여     
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 세액계산근거 - 세액공제(외국납부) - 법정이외 퇴직급여
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 결정세액 - 퇴직소득세 - 법정퇴직급여   
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 결정세액 - 퇴직소득세 - 법정이외 퇴직급여   
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 결정세액 - 퇴직주민세
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // 세후 지급액
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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



                if (vCountRow > 0)
                {
                    for (int vRow = 0; vRow < vCountRow; vRow++)
                    {
                        vGDColumnIndex1 = pGrid_ETC_ALLOWANCE.GetColumnToIndex("ALLOWANCE_DESC");
                        vXLColumnIndex1 = 16;

                        vGDColumnIndex2 = pGrid_ETC_ALLOWANCE.GetColumnToIndex("ALLOWANCE_AMOUNT");
                        vXLColumnIndex2 = 36;

                        vObject = pGrid_ETC_ALLOWANCE.GetCellValue(vRow, vGDColumnIndex1);
                        IsConvert = IsConvertString(vObject, out vConvertString);
                        if (IsConvert == true)
                        {
                            vConvertString = string.Format("{0}", vConvertString);
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex1, vConvertString);
                        }
                        else
                        {
                            vConvertString = string.Empty;
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex1, vConvertString);
                        }

                        vObject = pGrid_ETC_ALLOWANCE.GetCellValue(vRow, vGDColumnIndex2);
                        IsConvert = IsConvertString(vObject, out vConvertString);
                        IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                        if (IsConvert == true)
                        {
                            vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex2, vConvertString);
                        }
                        else
                        {
                            vConvertString = string.Empty;
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex2, vConvertString);
                        }
                        vXLine = vXLine + 1;
                    }
                }

                // 공제 및 가산항목(기타공제 사유)    
                //vGDColumnIndex = pGDColumn[74];
                //vXLColumnIndex = pXLColumn[74];
                //vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertString(vObject, out vConvertString);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0}", vConvertString);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                // 공제 및 가산항목(기타공제)  
                //vGDColumnIndex = pGDColumn[75];
                //vXLColumnIndex = pXLColumn[75];
                //vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                //-------------------------------------------------------------------
                vXLine = 50;
                //-------------------------------------------------------------------

                // 차인지급액
                vGDColumnIndex = pGDColumn[76];
                vXLColumnIndex = pXLColumn[76];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = 53;
                //-------------------------------------------------------------------

                // 은행
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 날짜
                vConvertString = string.Format("{0}", iDate.ISGetDate().ToShortDateString());
                mPrinting.XLSetCell(vXLine, 34, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                // 영수인
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

        private int XLLine2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_ETC_ALLOWANCE, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pCourse)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;
            

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                int vCountRow = pGrid_ETC_ALLOWANCE.RowCount; //pGrid_ETC_ALLOWANCE 그리드의 총 행수

                mPrinting.XLActiveSheet("Destination");

               
                // 항목 
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }



                // 마지막달 -3 
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 마지막달 -2
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 마지막달 -1
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 마지막달
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //(합계) 
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

        #region ----- Excel Write RetireAdjustment  Method ----

        public int WriteRetireAdjustment(string pPrint_Type, string pSaveFileName, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_ETC_ALLOWANCE)
        {
            
            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn;
            int[] vXLColumn;

            int vTotalRow = pGrid_RETIRE_ADJUSTMENT.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 0;

            int vSecondPrinting = 9; //1인당 3페이지이므로, 3*10=30번째에 인쇄
            int vCountPrinting = 0;

            SetArray1(pGrid_RETIRE_ADJUSTMENT, out vGDColumn, out vXLColumn);

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;
                pGrid_RETIRE_ADJUSTMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                if (isOpen == true)
                {
                    vCountPrinting++;

                    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                    vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                    pGrid_RETIRE_ADJUSTMENT.CurrentCellMoveTo(vRow, 0);
                    pGrid_RETIRE_ADJUSTMENT.Focus();
                    pGrid_RETIRE_ADJUSTMENT.CurrentCellActivate(vRow, 0);

                    // 퇴직금 정산 내역
                    vPrintingLine = XLLine1(pGrid_RETIRE_ADJUSTMENT, pGrid_ETC_ALLOWANCE, vRow, vPrintingLine, vGDColumn, vXLColumn, "SRC_TAB1");

                    if (vSecondPrinting < vCountPrinting)
                    {
                        if (pPrint_Type == "FILE")
                        {
                            ////파일 저장
                            SAVE(pSaveFileName);
                        }
                        else
                        {
                            Printing(1, vSecondPrinting);
                        }

                        mPrinting.XLOpenFileClose();
                        isOpen = XLFileOpen();

                        vCountPrinting = 0;
                        vPrintingLine = 1;
                        mCopyLineSUM = 1;
                    }
                    else if (vTotalRow == vRowCount)
                    {
                        if (pPrint_Type == "FILE")
                        {
                            ////파일 저장
                            SAVE(pSaveFileName);
                        }
                        else
                        {
                            Printing(1, vCountPrinting);
                            //PreView(1, vCountPrinting);       
                        }
                    }
                }
            }
            
            mPrinting.XLOpenFileClose();

            return mPageNumber;
        }

        public int WriteRetireAdjustment_NFK(string pPrint_Type, string pSaveFileName, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_ETC_ALLOWANCE)
        {
            mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치, 복사 행 누적
            mIncrementCopyMAX = 85; //복사되어질 행의 범위
           
            mCopyColumnSTART = 1;    //복사되어  진 행 누적 수
            mCopyColumnEND = 43;     //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn;
            int[] vXLColumn;

            int vTotalRow = pGrid_ETC_ALLOWANCE.RowCount;
            int vH_Row = 0;//  pGrid_RETIRE_ADJUSTMENT.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 30;

            int vSecondPrinting = 11; //1인당 3페이지이므로, 3*10=30번째에 인쇄
            int vCountPrinting = 0;

            SetArray2(pGrid_RETIRE_ADJUSTMENT, pGrid_ETC_ALLOWANCE, out vGDColumn, out vXLColumn);

            HeaderWrite(pGrid_RETIRE_ADJUSTMENT, vH_Row, vGDColumn, vXLColumn);

            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;
                pGrid_ETC_ALLOWANCE.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                if (isOpen == true)
                {
                    vCountPrinting++;

                    pGrid_ETC_ALLOWANCE.CurrentCellMoveTo(vRow, 0);
                    pGrid_ETC_ALLOWANCE.Focus();
                    pGrid_ETC_ALLOWANCE.CurrentCellActivate(vRow, 0);

                    // 퇴직금 정산 내역
                    vPrintingLine = XLLine2(pGrid_ETC_ALLOWANCE, vRow, vPrintingLine, vGDColumn, vXLColumn, "SRC_TAB1");

                    if (vTotalRow == vRowCount)
                    {
                        if (pPrint_Type == "FILE")
                        {
                            DeleteSheet();
                            ////파일 저장
                            SAVE(pSaveFileName);
                        }
                        else
                        {
                            DeleteSheet();
                            Printing(1, vCountPrinting);
                            //PreView(1, vCountPrinting);       
                        }
                    }
                }
            }

            mPrinting.XLOpenFileClose();

            return mPageNumber;
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

 
            pPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet("Destination");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber++; //페이지 번호

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        public void DeleteSheet()
        {
            mPrinting.XLDeleteSheet("SourceTab1");
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

        public void SAVE(string pSaveFileName)
        {
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder.ToString(), vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
        
        #endregion;
    }
}