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

namespace HRMF0633
{
    public class XLPrinting2
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
        private int mIncrementCopyMAX = 71; //복사되어질 행의 범위

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

        public XLPrinting2(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
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

        #region ----- SetArray1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[162];
            pXLColumn = new int[162];

            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[0]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TYPE_1");           // 거주 구분(거주자1)       
            pGDColumn[1]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TYPE_2");           // 거주 구분(거주자2)       
            pGDColumn[2]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONALITY_TYPE_1");        // 내외국인 구분(내국인1)   
            pGDColumn[3]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONALITY_TYPE_9");        // 내외국인 구분(외국인9)      
            pGDColumn[4]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATION_DESC");               // 거주지국
            pGDColumn[5]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATION_ISO_CODE");           // 거주지국코드
            pGDColumn[6]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OWNER_TYPE");                // 징수의무자구분
            pGDColumn[7]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("VAT_NUMBER");                // 사업자등록번호                                       
            pGDColumn[8]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CORP_NAME");                 // 법인명(상호)                                         
            pGDColumn[9]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRESIDENT_NAME");            // 대표자(성명)                                         
            pGDColumn[10]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ORG_ADDRESS");               // 소재지(주소)   
                                      
            pGDColumn[11]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NAME");                      // 성명                                                 
            pGDColumn[12]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPRE_NUM");                 // 주민번호                                             
            pGDColumn[13]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADDRESS");                   // 주소                                                 
            pGDColumn[14]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("START_RETIRE_DATE");         // 귀속연도 시작 일자                                   
            pGDColumn[15]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_NAME");               // 퇴직사유                                             
            pGDColumn[16]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LAST_RETIRE_DATE");          // 귀속연도 마지막 일자(퇴직일자)                       
            pGDColumn[17]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TAX1");               // 퇴직세액공제적용1                                    
            pGDColumn[18]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TAX2");               // 퇴직세액공제적용2                                    
            pGDColumn[19]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TAX3");               // 퇴직세액공제적용3                                    
            pGDColumn[20]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TAX4");               // 퇴직세액공제적용 합계   
                             
            pGDColumn[21]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME1");           // 근무처명1                                            
            pGDColumn[22]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME2");           // 근무처명2                                            
            pGDColumn[23]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME3");           // 근무처명3                                            
            pGDColumn[24]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME4");           // 근무처명 합계                                        
            pGDColumn[25]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER1");          // 사업자등록번호1                                      
            pGDColumn[26]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER2");          // 사업자등록번호2                                      
            pGDColumn[27]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER3");          // 사업자등록번호3                                      
            pGDColumn[28]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER4");          // 사업자등록번호 합계                                  
            pGDColumn[29]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TOTAL_AMOUNT1");      // 퇴직급여1                                            
            pGDColumn[30]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TOTAL_AMOUNT2");      // 퇴직급여2          
                                  
            pGDColumn[31]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TOTAL_AMOUNT3");      // 퇴직급여3                                            
            pGDColumn[32]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TOTAL_AMOUNT4");      // 퇴직급여 합계               
            pGDColumn[33]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT1");          // 명예퇴직수당(추가퇴직금)1                            
            pGDColumn[34]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT2");          // 명예퇴직수당(추가퇴직금)2                            
            pGDColumn[35]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT3");          // 명예퇴직수당(추가퇴직금)3                            
            pGDColumn[36]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT4");          // 명예퇴직수당(추가퇴직금) 합계                        
            pGDColumn[37]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT1");            // 퇴직연금일시금1                                      
            pGDColumn[38]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT2");            // 퇴직연금일시금2                                      
            pGDColumn[39]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT3");            // 퇴직연금일시금3                                      
            pGDColumn[40]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT4");            // 퇴직연금일시금 합계   

            pGDColumn[41]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_HONORARY_AMOUNT1");   // 명예퇴직수당 퇴직연금일시금  
            pGDColumn[42]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_HONORARY_AMOUNT2");   // 명예퇴직수당 퇴직연금일시금  
            pGDColumn[43]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_HONORARY_AMOUNT3");   // 명예퇴직수당 퇴직연금일시금  
            pGDColumn[44]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_HONORARY_AMOUNT4");   // 명예퇴직수당 퇴직연금일시금 합계  
            pGDColumn[45]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL1");                    // 법정계1                                                  
            pGDColumn[46]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL2");                    // 법정계2                                                  
            pGDColumn[47]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL3");                    // 법정계3                                                  
            pGDColumn[48]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL4");                    // 법정계 합계 
            pGDColumn[49] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_HONORARY_AMOUNT1");     // 법정외계1                                                  
            pGDColumn[50] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_HONORARY_AMOUNT2");     // 법정외계2  
                                                
            pGDColumn[51] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_HONORARY_AMOUNT3");     // 법정외계3                                                  
            pGDColumn[52] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_HONORARY_AMOUNT4");     // 법정외계 합계 
            pGDColumn[53]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NON_TAX1");                  // 비과세소득1                                          
            pGDColumn[54]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NON_TAX2");                  // 비과세소득2                                          
            pGDColumn[55]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NON_TAX3");                  // 비과세소득3                                          
            pGDColumn[56]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NON_TAX4");                  // 비과세소득 합계                                      
            pGDColumn[57]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_RECEIPTS1");           // 총수령액1                                            
            pGDColumn[58]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPAY_TOTAL_AMOUNT1");       // 원리금 합계액1               
            pGDColumn[59]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_MONEY_DUE1");         // 소득자 불입액1                                       
            pGDColumn[60]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANNUITY_DED1");       // 퇴직연금 소득공제액1                                 

            pGDColumn[61]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANNUITY_LUMP_SUM1");  // 퇴직연금일시금1                                      
            pGDColumn[62]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_RECEIPTS2");           // 총수령액2                                            
            pGDColumn[63]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPAY_TOTAL_AMOUNT2");       // 원리금 합계액2                                       
            pGDColumn[64]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_MONEY_DUE2");         // 소득자 불입액2                                       
            pGDColumn[65]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANNUITY_DED2");       // 퇴직연금 소득공제액2                                 
            pGDColumn[66]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANNUITY_LUMP_SUM2");  // 퇴직연금일시금2                                      
            pGDColumn[67]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANN_LUMP_SUM_E1");    // 퇴직연금일시금 지급예상액, 이연금액                  
            pGDColumn[68]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_LUMP_SUM1");           // 총일시금                                  
            pGDColumn[69]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECEIVE_RETIRE_PAY1");       // 수령가능퇴직급여액                                   
            pGDColumn[70]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_RETIRE_ANN_DED1");        // 환산퇴직소득공제
                                     
            pGDColumn[71]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_RETIRE_ANN_STANDARD1");   // 환산퇴직소득과세표준                                 
            pGDColumn[72]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_YEARLY_ANN_STANDARD1");   // 환산연평균 과세표준                                  
            pGDColumn[73]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_YEARLY_TAX_AMOUNT1");     // 환산 연평균 산출세액                                 
            pGDColumn[74]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANN_LUMP_SUM_E2");    // 퇴직연금일시금 지급예상액, 이연금액                  
            pGDColumn[75]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANN_LUMP_SUM_E3");    // 퇴직연금일시금 지급예상액, 이연금액                  
            pGDColumn[76]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_LUMP_SUM2");           // 총일시금                                             
            pGDColumn[77]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECEIVE_RETIRE_PAY2");       // 수령가능퇴직급여액               
            pGDColumn[78]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_RETIRE_ANN_DED2");        // 환산퇴직소득공제          
            pGDColumn[79]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_RETIRE_ANN_STANDARD2");   // 환산퇴직소득과세표준                                 
            pGDColumn[80]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_YEARLY_ANN_STANDARD2");   // 환산연평균 과세표준                                  

            pGDColumn[81]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_YEARLY_TAX_AMOUNT2");     // 환산 연평균 산출세액                                 
            pGDColumn[82]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANN_LUMP_SUM_E4");    // 퇴직연금일시금 지급예상액, 이연금액                  
            pGDColumn[83] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_DATE_FR1");           // 입사일(정산시작일)1     [ 법정 퇴직급여(주(현)근무지) ]                             
            pGDColumn[84]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_DATE1");              // 퇴사일1                                              
            pGDColumn[85]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_MONTH1");               // 근속월수1                                            
            pGDColumn[86]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EXCEPT_MONTH1");             // 제외월수1                                            
            pGDColumn[87]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_YEAR1");                // 근속연수1                                            
            pGDColumn[88] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ORI_JOIN_DATE2");            // 입사일2     [ 법정 외 퇴직급여(주(현)근무지) ]                                        
            pGDColumn[89]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_DATE2");              // 퇴사일2                                              
            pGDColumn[90]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_MONTH2");               // 근속월수2                                            

            pGDColumn[91]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EXCEPT_MONTH2");             // 제외월수2                                            
            pGDColumn[92]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_YEAR2");                // 근속연수2                                            
            pGDColumn[93] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_RETIRE_DATE_FR1");       // 입사일(정산시작일)1    [ 법정 퇴직급여(종(전)근무지) ]                              
            pGDColumn[94]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_RETIRE_DATE1");          // 퇴사일1                                              
            pGDColumn[95]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LONG_MONTH1");           // 근속월수1                                            
            pGDColumn[96]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_EXCEPT_MONTH1");         // 제외월수1                                            
            pGDColumn[97] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_ORI_JOIN_DATE2");        // 입사일2                 [ 법정 외 퇴직급여(종(전)근무지) ]                                          
            pGDColumn[98]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_RETIRE_DATE2");          // 퇴사일2                                              
            pGDColumn[99]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LONG_MONTH2");           // 근속월수2                                            
            pGDColumn[100]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_EXCEPT_MONTH2");         // 제외월수2  

            pGDColumn[101]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LONG_YEAR1");            // 중복월수1             [ 법정 퇴직급여(종(전)근무지) ]                    
            pGDColumn[102]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LONG_YEAR2");            // 중복월수2             [ 법정 외 퇴직급여(종(전)근무지) ]     
            pGDColumn[103]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT");             // 퇴직급여액                                           
            pGDColumn[104]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT");           // 명예퇴직금                                           
            pGDColumn[105]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_RETIRE_AMOUNT");       // 퇴직급여액 합계                                      
            pGDColumn[106]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DED_SUM_AMOUNT");            // 퇴직소득공제 - 계                                    
            pGDColumn[107]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_DED_SUM_AMOUNT");          // 퇴직소득공제 - 계                                    
            pGDColumn[108]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_INCOME_DED_AMOUNT");   // 퇴직소득공제 합계                                    /
            pGDColumn[109]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_STD_AMOUNT");            // 퇴직소득과세표준 - 법정퇴직급여                      
            pGDColumn[110]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_TAX_STD_AMOUNT");          // 퇴직소득과세표준 - 법정이외 퇴직급여                 

            pGDColumn[111] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_TAX_STD_AMOUNT");      // 퇴직소득과세표준 합계.                              
            pGDColumn[112] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("AVG_TAX_STD_AMOUNT");        // 세액계산근거 - 연평균과세표준 - 법정퇴직급여         
            pGDColumn[113] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_AVG_TAX_STD_AMOUNT");      // 세액계산근거 - 연평균과세표준 - 법정이외 퇴직급여    
            pGDColumn[114] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AVG_TAX_STD_AMOUNT");   // 연평균과세표준 합계.                   
            pGDColumn[115] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("AVG_COMP_TAX_AMOUNT");       // 세액계산근거 - 연평균산출세액 - 법정퇴직급여         
            pGDColumn[116] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_AVG_COMP_TAX_AMOUNT");     // 세액계산근거 - 연평균산출세액 - 법정이외 퇴직급여    
            pGDColumn[117] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AVG_COMP_TAX_AMOUNT"); // 연평균산출세액 합계                                  
            pGDColumn[118] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("COMP_TAX_AMOUNT");           // 세액계산근거 - 산출세액 - 법정퇴직급여               
            pGDColumn[119] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_COMP_TAX_AMOUNT");         // 세액계산근거 - 산출세액 - 법정이외 퇴직급여          
            pGDColumn[120] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_COMP_TAX_AMOUNT");     // 산출세액 합계                                        

            pGDColumn[121] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_AMOUNT");            // 세액계산근거 - 세액공제(외국납부) - 법정퇴직급여     
            pGDColumn[122] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_TAX_DED_AMOUNT");          // 세액계산근거 - 세액공제(외국납부) - 법정이외 퇴직급여
            pGDColumn[123] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_TAX_DED_AMOUNT");      // 세액공제 합계                                        
            pGDColumn[124] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT");         // 결정세액 - 퇴직소득세 - 법정퇴직급여                 
            pGDColumn[125] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_INCOME_TAX_AMOUNT");       // 결정세액 - 퇴직소득세 - 법정이외 퇴직급여            
            pGDColumn[126] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_INCOME_TAX_AMOUNT");   // 결정세액 합계                                        
            pGDColumn[127] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_INCOME_TAX_AMOUNT1");  // 결정세액 - 소득세 합계                               
            pGDColumn[128] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TAX_AMOUNT");       // 결정세액 - 주민세 합계                               
            pGDColumn[129] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_SP_TAX_AMOUNT");       // 농어촌 특별세 합계                                   
            pGDColumn[130] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_INCOME_TAX_AMOUNT2");  // 결정세액 합계                                        

            pGDColumn[131] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_INCOME_AMOUNT");         // 종(전)근무지 기납부세액 | 소득세                     
            pGDColumn[132] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LOCAL_AMOUNT");          // 종(전)근무지 기납부세액 | 주민세                     
            pGDColumn[133] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_SP_TAX_AMOUNT");         // 종(전)근무지 기납부세액 | 농특세                     
            pGDColumn[134] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_TOTAL");                 // 종(전)근무지 기납부세액 | 합계                       
            pGDColumn[135] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MINUS_INCOME_AMOUNT");       // 차감원천징수세액 | 소득세                            
            pGDColumn[136] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MINUS_LOCAL_AMOUNT");        // 차감원천징수세액 | 주민세                            
            pGDColumn[137] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MINUS_SP_TAX_AMOUNT");       // 차감원천징수세액 | 농특세                            
            pGDColumn[138] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_MINUS_TAX");           // 차감원천징수세액 합계


            pGDColumn[139] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRINT_DATE");                // 출력날짜
            pGDColumn[140] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRINT_OWNER");               // 징수의무자
            pGDColumn[141] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OFFICER_FLAG");              // 임원여부
            pGDColumn[142] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_1");           // 정년퇴직
            pGDColumn[143] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_2");           // 정리해고
            pGDColumn[144] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_3");           // 자발적퇴직
            pGDColumn[145] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_4");           // 임원퇴직
            pGDColumn[146] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_5");           // 중간정산
            pGDColumn[147] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_6");           // 기타
            pGDColumn[148] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_MONTH1");                // 법정퇴직급여 가산월수

            pGDColumn[149] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_MONTH2");                // 법정 외 퇴직급여 가산월수.
            pGDColumn[150] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIREMENT_PENSION_ACCOUNT");// 퇴직연금 계좌번호.
            pGDColumn[151] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DEFERRAL_TAX_AMT");          // 과세이연 금액.
            pGDColumn[152] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_RECEIPT_RETIRE_PAY_AMT");// 기수령한 퇴직급여액.
            pGDColumn[153] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RP_CORP_DESC");              // 퇴직연금 사업자명.
            pGDColumn[154] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RP_TAX_REG_NUM");            // 사업자번호
            pGDColumn[155] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RP_ACCOUNT_NUM");            // 계좌번호
            pGDColumn[156] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TRANS_RETIRE_AMOUNT");       // 입금(이체) 법정퇴직급여
            pGDColumn[157] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TRANS_HONORARY_AMOUNT");     // 입금(이체) 법정퇴직급여이외
            pGDColumn[158] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TRANS_DATE");                // 입금(이체) 일

            pGDColumn[159] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EXPIRE_DATE");               // 만기일
            pGDColumn[160] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TRANS_INCOME_TAX_AMOUNT");   // 과세이연 결정세액 합계


            //---------------------------------------------------------------------------------------------------------------------

            pXLColumn[0]   = 37;   // 거주 구분(거주자1)    
            pXLColumn[1]   = 42;   // 거주 구분(거주자2)      
            pXLColumn[2]   = 37;   // 내외국인 구분(내국인1)                                    
            pXLColumn[3]   = 42;   // 내외국인 구분(외국인9)                              
            pXLColumn[4]   = 31;   // 거주지국                                          
            pXLColumn[5]   = 39;   // 거주지국코드                                          
            pXLColumn[6]   = 33;   // 징수의무자구분                                                  
            pXLColumn[7]   = 11;   // 사업자등록번호                                              
            pXLColumn[8]   = 27;   // 법인명(상호)                                                   
            pXLColumn[9]   = 39;   // 대표자(성명)                                  
            pXLColumn[10]  = 27;   // 소재지(주소)   

            pXLColumn[11]  = 11;   // 성명                      
            pXLColumn[12]  = 27;   // 주민번호                                     
            pXLColumn[13]  = 11;   // 주소                                     
            pXLColumn[14]  = 11;   // 귀속연도 시작 일자                                    
            pXLColumn[15]  = 27;   // 퇴직사유                                  안씀                        
            pXLColumn[16]  = 11;   // 귀속연도 마지막 일자(퇴직일자)                                              
            pXLColumn[17]  = 20;   // 퇴직세액공제적용1                         안씀                      
            pXLColumn[18]  = 28;   // 퇴직세액공제적용2                         안씀                
            pXLColumn[19]  = 36;   // 퇴직세액공제적용3                         안씀               
            pXLColumn[20]  = 12;   // 퇴직세액공제적용 합계                     안씀

            pXLColumn[21]  = 12;   // 근무처명1                                       
            pXLColumn[22]  = 20;   // 근무처명2                                       
            pXLColumn[23]  = 28;   // 근무처명3                                   
            pXLColumn[24]  = 36;   // 근무처명 합계                             안씀                                         
            pXLColumn[25]  = 12;   // 사업자등록번호1                                             
            pXLColumn[26]  = 20;   // 사업자등록번호2                                             
            pXLColumn[27]  = 28;   // 사업자등록번호3                                          
            pXLColumn[28]  = 36;   // 사업자등록번호 합계                       안씀                          
            pXLColumn[29]  = 12;   // 퇴직급여1                       
            pXLColumn[30]  = 20;   // 퇴직급여2 

            pXLColumn[31] = 28;   // 퇴직급여3                        
            pXLColumn[32] = 36;   // 퇴직급여 합계                                      
            pXLColumn[33] = 16;   // 명예퇴직수당(추가퇴직금)1                                             
            pXLColumn[34] = 24;   // 명예퇴직수당(추가퇴직금)2 
            pXLColumn[35] = 32;   // 명예퇴직수당(추가퇴직금)3
            pXLColumn[36] = 36;   // 명예퇴직수당(추가퇴직금) 합계       
            pXLColumn[37] = 12;   // 퇴직연금일시금1                                                   
            pXLColumn[38] = 20;   // 퇴직연금일시금2                                                   
            pXLColumn[39] = 28;   // 퇴직연금일시금3                                               
            pXLColumn[40] = 36;   // 퇴직연금일시금 합계    

            pXLColumn[41]  = 16;   // 명예퇴직수당 퇴직연금일시금                                      
            pXLColumn[42]  = 24;   // 명예퇴직수당 퇴직연금일시금                                 
            pXLColumn[43]  = 32;   // 명예퇴직수당 퇴직연금일시금                           
            pXLColumn[44]  = 36;   // 명예퇴직수당 퇴직연금일시금 합계  
            pXLColumn[45]  = 12;   // 법정계1                                        
            pXLColumn[46]  = 20;   // 법정계2                                    
            pXLColumn[47]  = 28;   // 법정계3                              
            pXLColumn[48]  = 36;   // 법정계 합계                                        
            pXLColumn[49]  = 16;   // 법정외계1                                             
            pXLColumn[50]  = 24;   // 법정외계2

            pXLColumn[51]  = 32;   // 법정외계3                                        
            pXLColumn[52]  = 36;   // 법정외계 합계                             
            pXLColumn[53]  = 12;   // 비과세소득1                                        
            pXLColumn[54]  = 20;   // 비과세소득2                  
            pXLColumn[55]  = 28;   // 비과세소득3                                              
            pXLColumn[56]  = 36;   // 비과세소득 합계                                    
            pXLColumn[57]  = 17;   // 총수령액1                                      
            pXLColumn[58]  = 22;   // 원리금 합계액1                                 
            pXLColumn[59]  = 27;   // 소득자 불입액1                             
            pXLColumn[60]  = 32;   // 퇴직연금 소득공제액1        

            pXLColumn[61]  = 37;   // 퇴직연금일시금1                   
            pXLColumn[62]  = 17;   // 총수령액2                   
            pXLColumn[63]  = 22;   // 원리금 합계액2                                              
            pXLColumn[64]  = 27;   // 소득자 불입액2                                    
            pXLColumn[65]  = 32;   // 퇴직연금 소득공제액2                                       
            pXLColumn[66]  = 37;   // 퇴직연금일시금2                                   
            pXLColumn[67]  = 9;   // 퇴직연금일시금 지급예상액, 이연금액                                 
            pXLColumn[68]  = 21;   // 총일시금                               
            pXLColumn[69]  = 25;   // 수령가능퇴직급여액               
            pXLColumn[70]  = 29;   // 환산퇴직소득공제         

            pXLColumn[71] = 33;   // 환산퇴직소득과세표준                                               
            pXLColumn[72] = 37;   // 환산연평균 과세표준                                             
            pXLColumn[73] = 41;   // 환산 연평균 산출세액                                            
            pXLColumn[74] = 9;   // 퇴직연금일시금 지급예상액, 이연금액                                        
            pXLColumn[75] = 9;   // 퇴직연금일시금 지급예상액, 이연금액                                             
            pXLColumn[76] = 21;   // 총일시금                                               
            pXLColumn[77] = 25;   // 수령가능퇴직급여액                                             
            pXLColumn[78] = 29;   // 환산퇴직소득공제                                             
            pXLColumn[79] = 33;   // 환산퇴직소득과세표준                                          
            pXLColumn[80] = 37;   // 환산연평균 과세표준
   
            pXLColumn[81] = 41;   // 환산 연평균 산출세액                                               
            pXLColumn[82] = 9;   // 퇴직연금일시금 지급예상액, 이연금액                                             
            pXLColumn[83] = 12;   // 입사일(정산시작일)1  
            pXLColumn[84] = 15;   // 퇴사일1                                               
            pXLColumn[85] = 18;   // 근속월수1                                                
            pXLColumn[86] = 20;   // 제외월수1                                             
            pXLColumn[87] = 26;   // 근속연수1                                             
            pXLColumn[88] = 28;   // 입사일2
            pXLColumn[89] = 31;   // 퇴사일2            
            pXLColumn[90] = 34;   // 근속월수2 

            pXLColumn[91]  = 36;   // 제외월수2                                            
            pXLColumn[92]  = 42;   // 근속연수2
            pXLColumn[93]  = 9;   // 입사일(정산시작일)1                                    
            pXLColumn[94]  = 12;   // 퇴사일1                                     
            pXLColumn[95]  = 15;   // 근속월수1                                      
            pXLColumn[96]  = 18;   // 제외월수1
            pXLColumn[97]  = 28;   // 입사일2                 
            pXLColumn[98]  = 31;   // 퇴사일2                          
            pXLColumn[99]  = 34;   // 근속월수2        
            pXLColumn[100] = 36;   // 제외월수2  

            pXLColumn[101] = 24;   // 중복월수1   ??                        
            pXLColumn[102] = 24;   // 중복월수2   ??      
            pXLColumn[103] = 12;   // 퇴직급여액     
            pXLColumn[104] = 25;   // 명예퇴직금  ??                                
            pXLColumn[105] = 38;   // 퇴직급여액 합계               
            pXLColumn[106] = 12;   // 퇴직소득공제 - 계           
            pXLColumn[107] = 25;   // 퇴직소득공제 - 계                                 
            pXLColumn[108] = 38;   // 퇴직소득공제 합계   
            pXLColumn[109] = 12;   // 퇴직소득과세표준 - 법정퇴직급여  
            pXLColumn[110] = 25;   // 퇴직소득과세표준 - 법정이외 퇴직급여

            pXLColumn[111] = 38;   // 퇴직소득과세표준 합계                
            pXLColumn[112] = 12;   // 세액계산근거 - 연평균과세표준 - 법정퇴직급여              
            pXLColumn[113] = 25;   // 세액계산근거 - 연평균과세표준 - 법정이외 퇴직급여                                        
            pXLColumn[114] = 38;   // 연평균과세표준 합계.                        
            pXLColumn[115] = 12;   // 세액계산근거 - 연평균산출세액 - 법정퇴직급여                     
            pXLColumn[116] = 25;   // 세액계산근거 - 연평균산출세액 - 법정이외 퇴직급여                                
            pXLColumn[117] = 38;   // 연평균산출세액 합계                                       
            pXLColumn[118] = 12;   // 세액계산근거 - 산출세액 - 법정퇴직급여                        
            pXLColumn[119] = 25;   // 세액계산근거 - 산출세액 - 법정이외 퇴직급여                     
            pXLColumn[120] = 28;   // 산출세액 합계                  

            pXLColumn[121] = 12;   // 세액계산근거 - 세액공제(외국납부) - 법정퇴직급여                  
            pXLColumn[122] = 25;   // 세액계산근거 - 세액공제(외국납부) - 법정이외 퇴직급여                  
            pXLColumn[123] = 28;   // 세액공제 합계                             
            //pXLColumn[124] = 27;   // 결정세액 - 퇴직소득세 - 법정퇴직급여                            
            //pXLColumn[125] = 35;   // 결정세액 - 퇴직소득세 - 법정이외 퇴직급여 
            //pXLColumn[126] = 29;   // 결정세액 합계 
            pXLColumn[127] = 12;   // 결정세액 - 소득세 합계
            pXLColumn[128] = 20;   // 결정세액 - 주민세 합계    
            pXLColumn[129] = 27;   // 농어촌 특별세 합계   
            pXLColumn[130] = 35;   // 결정세액 합계  

            pXLColumn[131] = 12;   // 종(전)근무지 기납부세액 | 소득세               
            pXLColumn[132] = 20;   // 종(전)근무지 기납부세액 | 주민세              
            pXLColumn[133] = 27;   // 종(전)근무지 기납부세액 | 농특세                                       
            pXLColumn[134] = 35;   // 종(전)근무지 기납부세액 | 합계                         
            pXLColumn[135] = 12;   // 차감원천징수세액 | 소득세                     
            pXLColumn[136] = 20;   // 차감원천징수세액 | 주민세                                  
            pXLColumn[137] = 27;   // 차감원천징수세액 | 농특세                                    
            pXLColumn[138] = 35;   // 차감원천징수세액 합계                  
            pXLColumn[139] = 35;   // 출력날짜                    
            pXLColumn[140] = 27;   // 징수의무자

            pXLColumn[141] = 39;   // 임원여부              
            pXLColumn[142] = 28;   // 정년퇴직          
            pXLColumn[143] = 34;   // 정리해고                                    
            pXLColumn[144] = 40;   // 자발적퇴직                         
            pXLColumn[145] = 28;   // 임원퇴직                     
            pXLColumn[146] = 34;   // 중간정산                                  
            pXLColumn[147] = 40;   // 기타
            pXLColumn[148] = 22;   // 법정퇴직급여 가산월수           
            pXLColumn[149] = 38;   // 법정 외 퇴직급여 가산월수.  
            pXLColumn[150] = 12;   // 퇴직연금 계좌번호.

            pXLColumn[151] = 13;   // 과세이연 금액.              
            pXLColumn[152] = 17;   // 기수령한 퇴직급여액.      
            pXLColumn[153] = 3;   // 퇴직연금 사업자명.                         
            pXLColumn[154] = 9;   // 사업자번호                  
            pXLColumn[155] = 15;   // 계좌번호          
            pXLColumn[156] = 21;   // 입금(이체) 법정퇴직급여                              
            pXLColumn[157] = 26;   // 입금(이체) 법정퇴직급여이외
            pXLColumn[158] = 31;   // 입금(이체) 일    
            pXLColumn[159] = 39;   // 만기일
            pXLColumn[160] = 39;   // 과세이연 결정세액 합계

        }

        #endregion;

        #region ----- SetArray2 ----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_2013, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[135];
            pXLColumn = new int[135];

            //--------------------------------------------------------------------------------------------------------------------

            // 우측 상단의 항목(거주구분, 내/외국인, 거주지국)
            pGDColumn[0] = pGrid_PRINT_2013.GetColumnToIndex("RESIDENT_TYPE_1");                           // 거주 구분(거주자1)       
            pGDColumn[1] = pGrid_PRINT_2013.GetColumnToIndex("RESIDENT_TYPE_2");                           // 거주 구분(거주자2)       
            pGDColumn[2] = pGrid_PRINT_2013.GetColumnToIndex("NATIONALITY_TYPE_1");                        // 내외국인 구분(내국인1)   
            pGDColumn[3] = pGrid_PRINT_2013.GetColumnToIndex("NATIONALITY_TYPE_9");                        // 내외국인 구분(외국인9)      
            pGDColumn[4] = pGrid_PRINT_2013.GetColumnToIndex("NATION_DESC");                               // 거주지국
            pGDColumn[5] = pGrid_PRINT_2013.GetColumnToIndex("NATION_ISO_CODE");                           // 거주지국코드
            pGDColumn[6] = pGrid_PRINT_2013.GetColumnToIndex("OWNER_TYPE");                                // 징수의무자구분
            
            // 징수의무자 인적사항
            pGDColumn[7] = pGrid_PRINT_2013.GetColumnToIndex("VAT_NUMBER");                                // 사업자등록번호                                       
            pGDColumn[8] = pGrid_PRINT_2013.GetColumnToIndex("CORP_NAME");                                 // 법인명(상호)                                         
            pGDColumn[9] = pGrid_PRINT_2013.GetColumnToIndex("PRESIDENT_NAME");                            // 대표자(성명)   
            pGDColumn[10] = pGrid_PRINT_2013.GetColumnToIndex("LEGAL_NUMBER");                             // 법인번호 
            pGDColumn[11] = pGrid_PRINT_2013.GetColumnToIndex("ORG_ADDRESS");                              // 소재지(주소)   

            //소득자 인적사항
            pGDColumn[12] = pGrid_PRINT_2013.GetColumnToIndex("NAME");                                     // 성명                                                 
            pGDColumn[13] = pGrid_PRINT_2013.GetColumnToIndex("REPRE_NUM");                                // 주민번호                                             
            pGDColumn[14] = pGrid_PRINT_2013.GetColumnToIndex("ADDRESS");                                  // 주소        
            pGDColumn[15] = pGrid_PRINT_2013.GetColumnToIndex("OFFICER_FLAG");                             // 임원여부
            pGDColumn[16] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_PENSION_REGIST_DATE");               // 확정급여형 퇴직연금 제도 가입일.    
            pGDColumn[17] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE");                              // 퇴직금 날짜.                 
            pGDColumn[18] = pGrid_PRINT_2013.GetColumnToIndex("START_RETIRE_DATE");                        // 귀속연도 시작 일자                               
            pGDColumn[19] = pGrid_PRINT_2013.GetColumnToIndex("LAST_RETIRE_DATE");                         // 귀속연도 마지막 일자(퇴직일자).                                 
            pGDColumn[20] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_1");                          // 정년퇴직                                 
            pGDColumn[21] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_2");                          // 정리해고  
            pGDColumn[22] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_3");                          // 자발적퇴직      
            pGDColumn[23] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_4");                          // 임원퇴직                                            
            pGDColumn[24] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_5");                          // 중간정산                                            
            pGDColumn[25] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_6");                          // 기타    

            //퇴직급여현황.
            pGDColumn[26] = pGrid_PRINT_2013.GetColumnToIndex("WORK_CORP_NAME1");                         // 근무처명 ( 중간지급)                  
            pGDColumn[27] = pGrid_PRINT_2013.GetColumnToIndex("WORK_CORP_NAME2");                         // 근무처명 (최종분)                                   
            pGDColumn[28] = pGrid_PRINT_2013.GetColumnToIndex("WORK_VAT_NUMBER1");                         // 사업자등록번호 ( 중간지급)                                    
            pGDColumn[29] = pGrid_PRINT_2013.GetColumnToIndex("WORK_VAT_NUMBER2");                         // 사업자등록번호 (최종분)                             
            pGDColumn[30] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_TOTAL_AMOUNT1");                     // 퇴직급여 (중간지급)
            pGDColumn[31] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_TOTAL_AMOUNT2");                     // 퇴직급여 (최종분)         
            pGDColumn[32] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_TOTAL_AMOUNT3");                     // 퇴직급여 (정산)
            pGDColumn[33] = pGrid_PRINT_2013.GetColumnToIndex("NON_TAX_TOTAL_AMOUNT1");                    // 비과세 퇴직급여 (중간지급)             
            pGDColumn[34] = pGrid_PRINT_2013.GetColumnToIndex("NON_TAX_TOTAL_AMOUNT2");                    // 비과세 퇴직급여 (최종분)                        
            pGDColumn[35] = pGrid_PRINT_2013.GetColumnToIndex("NON_TAX_TOTAL_AMOUNT3");                    // 과세 퇴직급여 (정산)
            pGDColumn[36] = pGrid_PRINT_2013.GetColumnToIndex("TAX_RETIRE_TOTAL_AMOUNT1");                 // 과세대상 퇴직급여 (중간지급)
            pGDColumn[37] = pGrid_PRINT_2013.GetColumnToIndex("TAX_RETIRE_TOTAL_AMOUNT2");                 // 과세대상 퇴직급 (최종분)                       
            pGDColumn[38] = pGrid_PRINT_2013.GetColumnToIndex("TAX_RETIRE_TOTAL_AMOUNT3");                 // 과세대상 퇴직급여 (정산)      

            //근속연수.
            pGDColumn[39] = pGrid_PRINT_2013.GetColumnToIndex("ORI_JOIN_DATE1");                           //입사일 ( 중간지급 근속연수)                              
            pGDColumn[40] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR1");                          // 기산일/정산시작일 ( 중간지급 근속연수)                                      
            pGDColumn[41] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE1");                             // 퇴사일 ( 중간지급 근속연수)
            pGDColumn[42] = pGrid_PRINT_2013.GetColumnToIndex("CLOSED_DATE1");                             // 지급일 ( 중간지급 근속연수)
            pGDColumn[43] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH1");                              // 근속월수 ( 중간지급 근속연수)
            pGDColumn[44] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH1");                            // 제외월수 ( 중간지급 근속연수)
            pGDColumn[45] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH1");                            // 가감월수 ( 중간지급 근속연수)                                    
            pGDColumn[46] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR1");                               // 근속연수 ( 중간지급 근속연수)


            pGDColumn[47] = pGrid_PRINT_2013.GetColumnToIndex("ORI_JOIN_DATE2");                           // 입사일(최종분)                                               
            pGDColumn[48] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR2");                          // 기산일/정산시작일 ( 최종분)
            pGDColumn[49] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE2");                             // 퇴사일 ( 최종분 )                                              
            pGDColumn[50] = pGrid_PRINT_2013.GetColumnToIndex("CLOSED_DATE2");                             // 지급일 ( 최종분 )
            pGDColumn[51] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH2");                              // 근속월수 ( 최종분)                                                   
            pGDColumn[52] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH2");                            // 법정외계 합계 
            pGDColumn[53] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH2");                            // 비과세소득1     
            pGDColumn[54] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR2");                               // 비과세소득2  


            pGDColumn[55] = pGrid_PRINT_2013.GetColumnToIndex("ORI_JOIN_DATE3");                           // 입사일(정산(합산))                                   
            pGDColumn[56] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR3");                          // 기산일/정산시작일 (정산(합산))                                   
            pGDColumn[57] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE3");                             // 퇴사일 (정산(합산))
            pGDColumn[58] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH3");                              // 근속월수 (정산(합산))
            pGDColumn[59] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH3");                            // 제외월수 (정산(합산))                                    
            pGDColumn[60] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH3");                            // 가감월수 (정산(합산))                             
            pGDColumn[61] = pGrid_PRINT_2013.GetColumnToIndex("PRE_LONG_YEAR3");                           // 중복월수 (정산(합산))                                  
            pGDColumn[62] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR3");                               // 근속연수 (정산(합산)) 

            pGDColumn[63] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR4");                          // 기산일 (2012.12.31 이전)                                  
            pGDColumn[64] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE4");                             // 퇴사일  (2012.12.31 이전)                            
            pGDColumn[65] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH4");                              // 근속월수 (2012.12.31 이전)     
            pGDColumn[66] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH4");                            // 제외월수 (2012.12.31 이전)
            pGDColumn[67] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH4");                            // 가감월수 (2012.12.31 이전)           
            pGDColumn[68] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR4");                               // 근속년수 (2012.12.31 이전)

            pGDColumn[69] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR5");                          // 기산일 ( 2013.01.01 이후)                                   
            pGDColumn[70] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE5");                             // 퇴사일  ( 2013.01.01 이후)
            pGDColumn[71] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH5");                              // 근속월수 ( 2013.01.01 이후)                         
            pGDColumn[72] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH5");                            // 제외월수 ( 2013.01.01 이후)   
            pGDColumn[73] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH5");                            // 가감월수 ( 2013.01.01 이후)
            pGDColumn[74] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR5");                               // 근속년수 ( 2013.01.01 이후)

            //퇴직소득과세표준계산
            pGDColumn[75] = pGrid_PRINT_2013.GetColumnToIndex("MID_RETIRE_AMOUNT");                        // 퇴직소득(중간지급)   
            pGDColumn[76] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_AMOUNT");                            // 퇴직소득(최종분)   
            pGDColumn[77] = pGrid_PRINT_2013.GetColumnToIndex("SUM_RETIRE_AMOUNT");                        // 퇴직소득(정산(합산)        
            pGDColumn[78] = pGrid_PRINT_2013.GetColumnToIndex("INCOME_DED_AMOUNT");                        // 퇴직소득정률공제     
            pGDColumn[79] = pGrid_PRINT_2013.GetColumnToIndex("LONG_DED_AMOUNT");                          // 근속연수공제                           
            pGDColumn[80] = pGrid_PRINT_2013.GetColumnToIndex("TAX_STD_AMOUNT");                           // 퇴직소득과세표준(27-28-29)    


            //퇴직소득세액계산
            pGDColumn[81] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_ANBON1");                           // 과세표준안분(2012.12.31이전)          
            pGDColumn[82] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_STD1");                             // 연평균과세표준(2012.12.31이전)
            pGDColumn[83] = pGrid_PRINT_2013.GetColumnToIndex("AVG_YEAR_COMP_TAX1");                       // 연평균산출세액(2012.12.31이전)                          
            pGDColumn[84] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX1");                                // 산출세액 (2012.12.31이전)

            pGDColumn[85] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_ANBON2");                           // 과세표준안분(2013.01.01이후 )                                      
            pGDColumn[86] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_STD2");                             // 과세표준(2013.01.01이후 )                                          
            pGDColumn[87] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_AVG_TAX_2");                         // 환산과세표준 (2013.01.01이후 )                          
            pGDColumn[88] = pGrid_PRINT_2013.GetColumnToIndex("CAHNGE_AVG_COMP_TAX2");                     // 환산산출세액 (2013.01.01이후 )                       
            pGDColumn[89] = pGrid_PRINT_2013.GetColumnToIndex("AVG_YEAR_COMP_TAX2");                       // 연평균산출세액 (2013.01.01이후 )                                         
            pGDColumn[90] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX2");                                // 산출세액 (2013.01.01이후 )

            pGDColumn[91] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_ANBON_SUM");                        // 과세표준안분(정산(합산))                                
            pGDColumn[92] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_STD_SUM");                          // 연평균과세표준(정산(합산))                                    
            pGDColumn[93] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_AVG_TAX_SUM");                       // 환산과세표준 
            pGDColumn[94] = pGrid_PRINT_2013.GetColumnToIndex("CAHNGE_AVG_COMP_TAX_SUM");                  // 환산산출세액액                                           
            pGDColumn[95] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX_AMOUNT_SUM");                      // 연평균산출세액                                           
            pGDColumn[96] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX_SUM");                             // 산출세액                                     
            pGDColumn[97] = pGrid_PRINT_2013.GetColumnToIndex("AHEAD_TAX");                                // 기납부세액                                
            pGDColumn[98] = pGrid_PRINT_2013.GetColumnToIndex("REPORT_TAX_SUM");                           // 신고대상세액 

            //이연퇴직소득세액계산
            pGDColumn[99] = pGrid_PRINT_2013.GetColumnToIndex("REPORT_TAX_SUM_39");                        // 신고대상세액 (39)                                  
            pGDColumn[100] = pGrid_PRINT_2013.GetColumnToIndex("RETIREMENT_CORP_NAME");                    // 연금계좌취급자
            pGDColumn[101] = pGrid_PRINT_2013.GetColumnToIndex("TAX_REG_NUM");                             // 사업자등록번호        
            pGDColumn[102] = pGrid_PRINT_2013.GetColumnToIndex("ACCOUNT_NUM");                             // 계좌번호
            pGDColumn[103] = pGrid_PRINT_2013.GetColumnToIndex("ISSUE_DATE");                              // 입금일                                           
            pGDColumn[104] = pGrid_PRINT_2013.GetColumnToIndex("TRANS_ACCOUNT_AMOUNT_1");                  // 계좌입금금액                                    
            pGDColumn[105] = pGrid_PRINT_2013.GetColumnToIndex("RETIREMENT_PAY");                          // 퇴직급여                                
            pGDColumn[106] = pGrid_PRINT_2013.GetColumnToIndex("TRANS_INCOME_TAX_AMOUNT");                 // 이연퇴직소득세        

            //납부명세
            pGDColumn[107] = pGrid_PRINT_2013.GetColumnToIndex("INCOME_TAX_AMOUNT");                       // 신고대상세액(소득세)                            
            pGDColumn[108] = pGrid_PRINT_2013.GetColumnToIndex("RESIDENT_TAX_AMOUNT");                     // 신고대상세액(지방소득세)
            pGDColumn[109] = pGrid_PRINT_2013.GetColumnToIndex("SP_TAX_AMOUNT");                           // 신고대상세액(농어촌특별세)
            pGDColumn[110] = pGrid_PRINT_2013.GetColumnToIndex("SUM_INCOME_RESIDENT");                     // 신고대상세액(계)

            pGDColumn[111] = pGrid_PRINT_2013.GetColumnToIndex("DEFERRED_INCOME_TAX");                     // 이연퇴직소득세(소득세)           
            pGDColumn[112] = pGrid_PRINT_2013.GetColumnToIndex("DEFERRED_RESIDENT_TAX");                   // 이연퇴직소득세(지방소득세)    
            pGDColumn[113] = pGrid_PRINT_2013.GetColumnToIndex("DEFERRED_SP_TAX");                         // 이연퇴직소득세(농어촌특별세)    
            pGDColumn[114] = pGrid_PRINT_2013.GetColumnToIndex("DEFERRED_SUM_INCOME_RESIDENT");            // 이연퇴직소득세(계)     

            pGDColumn[115] = pGrid_PRINT_2013.GetColumnToIndex("DEDUCTION_INCOME_TAX");                    // 차감원청징수세액(소득세)
            pGDColumn[116] = pGrid_PRINT_2013.GetColumnToIndex("DEDUCTION_RESIDENT_TAX");                  // 차감원청징수세액(지방소득세)  
            pGDColumn[117] = pGrid_PRINT_2013.GetColumnToIndex("DEDUCTION_SP_TAX");                        // 차감원청징수세액(농어촌특별세)  
            pGDColumn[118] = pGrid_PRINT_2013.GetColumnToIndex("DEDUCTION_SUM_INCOME_RESIDENT");           // 차감원청징수세액(계)  

            //하단프린트.
            pGDColumn[119] = pGrid_PRINT_2013.GetColumnToIndex("PRINT_DATE");                              // 출력일자 
            pGDColumn[120] = pGrid_PRINT_2013.GetColumnToIndex("PRINT_OWNER");                             // 대표이사                               
            pGDColumn[121] = pGrid_PRINT_2013.GetColumnToIndex("FINAL_RETIRE_DATE");                       // 정산지급

            //2016 추가분
            pGDColumn[122] = pGrid_PRINT_2013.GetColumnToIndex("SUM_RETIRE_AMOUNT_3");                     // 31.퇴직소득(정산(합산))
            pGDColumn[123] = pGrid_PRINT_2013.GetColumnToIndex("LONG_DED_AMOUNT_3");                       // 32.근속연수공제
            pGDColumn[124] = pGrid_PRINT_2013.GetColumnToIndex("CHG_STD_AMOUNT_3");                        // 33.환산급여
            pGDColumn[125] = pGrid_PRINT_2013.GetColumnToIndex("CHG_DED_AMOUNT_3");                        // 34.환산급여별공제
            pGDColumn[126] = pGrid_PRINT_2013.GetColumnToIndex("CHG_TAX_STD_AMOUNT_3");                    // 35.퇴직소득과세표준
            pGDColumn[127] = pGrid_PRINT_2013.GetColumnToIndex("CHG_COMP_TAX_AMOUNT_3");                   // 42.환산산출세액
            pGDColumn[128] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX_AMOUNT_3");                       // 43.산출세액
            pGDColumn[129] = pGrid_PRINT_2013.GetColumnToIndex("ADJUSTMENT_YYYY_3");                       // 44.퇴직일이 속하는 과세연도
            pGDColumn[130] = pGrid_PRINT_2013.GetColumnToIndex("REAL_COMP_TAX_AMOUNT_3");                  // 45.특례적용산출세액
            pGDColumn[131] = pGrid_PRINT_2013.GetColumnToIndex("PREPAID_TAX_AMOUNT_3");                    // 46.기납부 세액
            pGDColumn[132] = pGrid_PRINT_2013.GetColumnToIndex("INCOME_TAX_AMOUNT_3");                     // 47.신고대상세액
            pGDColumn[133] = pGrid_PRINT_2013.GetColumnToIndex("TRANS_ACCOUNT_SUM_AMOUNT");                // 연금계좌 입금합계 
            pGDColumn[134] = pGrid_PRINT_2013.GetColumnToIndex("TAX_OFFICE_NAME");                         // 세무서장
            

            //---------------------------------------------------------------------------------------------------------------------

            pXLColumn[0] = 37;   // 거주 구분(거주자1)    
            pXLColumn[1] = 42;   // 거주 구분(거주자2)   
   
            pXLColumn[2] = 37;   // 내외국인 구분(내국인1)                                    
            pXLColumn[3] = 42;   // 내외국인 구분(외국인9)   
                           
            pXLColumn[4] = 31;   // 거주지국                                          
            pXLColumn[5] = 39;   // 거주지국코드          
                                
            pXLColumn[6] = 33;   // 징수의무자구분  
            
            //---------------------------------------
            pXLColumn[7] = 11;   // 사업자등록번호                                              
            pXLColumn[8] = 27;   // 법인명(상호)                                                   
            pXLColumn[9] = 39;   // 대표자(성명)

            pXLColumn[10] = 11;   // 법인번호 
            pXLColumn[11] = 27;   // 소재지(주소)   

            //-----------------------------------------
            pXLColumn[12] = 11;   // 성명                      
            pXLColumn[13] = 27;   // 주민번호     
                                
            pXLColumn[14] = 11;   // 주소                                     
            pXLColumn[15] = 39;   // 임원여부

            pXLColumn[16] = 11;   // 확정급여형 퇴직연금 제도 가입일
            pXLColumn[17] = 36;   // 퇴직금 날짜                    
                         
            pXLColumn[18] = 11;   // 귀속연도 시작 일자            
            pXLColumn[19] = 11;   // 귀속연도 마지막 일자(퇴직일자).                
            pXLColumn[20] = 26;   // 정년퇴직                
            pXLColumn[21] = 32;   // 정리해고  
            pXLColumn[22] = 38;   // 자발적퇴직                                        
            pXLColumn[23] = 26;   // 임원퇴직                                      
            pXLColumn[24] = 32;   // 중간정산                                    
            pXLColumn[25] = 38;   // 기타
            //-----------------------------------------

            pXLColumn[26] = 14;   // 근무처명 ( 중간지급)                                             
            pXLColumn[27] = 24;   // 근무처명 ( 중간지급)         
                                    
            pXLColumn[28] = 14;   // 사업자등록번호 ( 중간지급)                                            
            pXLColumn[29] = 24;   // 사업자등록번호 (최종분)  
           
            pXLColumn[30] = 14;   // 퇴직급여 (중간지급)                   
            pXLColumn[31] = 24;   // 퇴직급여 (최종분)                        
            pXLColumn[32] = 34;   // 퇴직급여 (정산)    

            pXLColumn[33] = 14;   // 비과세 퇴직급여 (중간지급)                                            
            pXLColumn[34] = 24;   // 비과세 퇴직급여 (최종분)      
            pXLColumn[35] = 34;   // 과세 퇴직급여 (정산)

            pXLColumn[36] = 14;   // 과세대상 퇴직급여 (중간지급) 
            pXLColumn[37] = 24;   // 과세대상 퇴직급 (최종분)                                                   
            pXLColumn[38] = 34;   // 과세대상 퇴직급여 (정산)      

            //-----------------------------------------

            pXLColumn[39] = 13;   // 입사일 ( 중간지급 근속연수)                                             
            pXLColumn[40] = 17;   // 기산일/정산시작일 ( 중간지급 근속연수)    
            pXLColumn[41] = 21;   // 퇴사일 ( 중간지급 근속연수)                                
            pXLColumn[42] = 25;   // 지급일 ( 중간지급 근속연수)                            
            pXLColumn[43] = 29;   // 근속월수 ( 중간지급 근속연수)                       
            pXLColumn[44] = 32;   // 제외월수 ( 중간지급 근속연수)  
            pXLColumn[45] = 35;   // 가산월수 ( 중간지급 근속연수)                                        
            pXLColumn[46] = 41;   // 근속연수 ( 중간지급 근속연수)  

            pXLColumn[47] = 13;   // 입사일( 최종분)                               
            pXLColumn[48] = 17;   // 기산일/정산시작일 ( 최종분)                              
            pXLColumn[49] = 21;   // 퇴사일 ( 최종분 )                                                
            pXLColumn[50] = 25;   // 지급일 ( 최종분 )
            pXLColumn[51] = 29;   // 근속월수 ( 최종분)                                            
            pXLColumn[52] = 32;   // 제외월수 ( 최종분)                            
            pXLColumn[53] = 35;   // 가산월수 ( 최종분)                                       
            pXLColumn[54] = 41;   // 근속연수 ( 최종분)  

            pXLColumn[55] = 13;   // 입사일(정산(합산))                                              
            pXLColumn[56] = 17;   // 기산일/정산시작일 (정산(합산))                                   
            pXLColumn[57] = 21;   // 퇴사일 (정산(합산))                                 
            pXLColumn[58] = 29;   // 근속월수 (정산(합산))                             
            pXLColumn[59] = 32;   // 제외월수 (정산(합산))                                
            pXLColumn[60] = 35;   // 가산월수 (정산(합산))       
            pXLColumn[61] = 38;   // 중복월수 (정산(합산))            
            pXLColumn[62] = 41;   // 근속연수 (정산(합산))    

            pXLColumn[63] = 17;   // 기산일 (2012.12.31 이전)                                      
            pXLColumn[64] = 21;   // 퇴사일  (2012.12.31 이전)                              
            pXLColumn[65] = 29;   // 근속월수 (2012.12.31 이전)                                  
            pXLColumn[66] = 32;   // 제외월수 (2012.12.31 이전)                             
            pXLColumn[67] = 35;   // 가산월수 (2012.12.31 이전)                                  
            pXLColumn[68] = 41;   // 근속년수 (2012.12.31 이전)

            pXLColumn[69] = 17;   // 기산일 ( 2013.01.01 이후)                   
            pXLColumn[70] = 21;   // 퇴사일  ( 2013.01.01 이후)         
            pXLColumn[71] = 29;   // 근속월수 ( 2013.01.01 이후)                                             
            pXLColumn[72] = 32;   // 제외월수 ( 2013.01.01 이후)                                
            pXLColumn[73] = 35;   // 가감월수 ( 2013.01.01 이후)                             
            pXLColumn[74] = 41;    // 근속년수 ( 2013.01.01 이후)    


            //-----------------------------------------

            pXLColumn[75] = 24;   // 34.퇴직소득(중간지급)                                            
            pXLColumn[76] = 24;   // 34.퇴직소득(최종분)                                             
            pXLColumn[77] = 24;   // 34.퇴직소득(정산(합산) 

            pXLColumn[78] = 24;   // 35.퇴직소득정률공제                                         
            pXLColumn[79] = 24;   // 36.근속연수공제                            
            pXLColumn[80] = 24;   // 37.퇴직소득과세표준(27-28-29)    


            //-----------------------------------------
            pXLColumn[81] = 24;   // 과세표준안분(2012.12.31이전)                                       
            pXLColumn[82] = 24;   // 연평균과세표준(2012.12.31이전)                                   
            pXLColumn[83] = 24;   // 연평균산출세액(2012.12.31이전) 
            pXLColumn[84] = 24;   // 산출세액 (2012.12.31이전)

            pXLColumn[85] = 30;   // 과세표준안분(2013.01.01이후 )                                               
            pXLColumn[86] = 30;   // 과세표준(2013.01.01이후 )                                     
            pXLColumn[87] = 30;   // 환산과세표준 (2013.01.01이후 )                                          
            pXLColumn[88] = 30;   // 환산산출세액 (2013.01.01이후 )  
            pXLColumn[89] = 30;   // 연평균산출세액 (2013.01.01이후 )            
            pXLColumn[90] = 30;   // 산출세액 (2013.01.01이후 )


            pXLColumn[91] = 36;   // 과세표준안분(정산(합산))                                             
            pXLColumn[92] = 36;   // 연평균과세표준(정산(합산))
            pXLColumn[93] = 36;   // 환산과세표준                             
            pXLColumn[94] = 36;   // 환산산출세액액                                        
            pXLColumn[95] = 36;   // 연평균산출세액                            
            pXLColumn[96] = 36;   // 산출세액 
            pXLColumn[97] = 24;   // 기납부세액             
            pXLColumn[98] = 24;   // 신고대상세액 


            //-----------------------------------------
            pXLColumn[99]  = 3;    // 신고대상세액 (39) 
            pXLColumn[100] = 7;    // 연금계좌취급자
            pXLColumn[101] = 13;   // 사업자등록번호             
            pXLColumn[102] = 18;   // 계좌번호
            pXLColumn[103] = 24;   // 입금일  
            pXLColumn[104] = 29;   // 계좌입금금액                                
            pXLColumn[105] = 34;   // 퇴직급여           
            pXLColumn[106] = 39;   // 이연퇴직소득세        


            //-----------------------------------------
            pXLColumn[107] = 13;   // 신고대상세액(소득세)                                
            pXLColumn[108] = 21;   // 신고대상세액(지방소득세)
            pXLColumn[109] = 28;   // 신고대상세액(농어촌특별세)
            pXLColumn[110] = 36;   // 신고대상세액(농어촌특별세)

            pXLColumn[111] = 13;   // 이연퇴직소득세(소득세)
            pXLColumn[112] = 21;   // 이연퇴직소득세(지방소득세)     
            pXLColumn[113] = 28;   // 이연퇴직소득세(농어촌특별세) 
            pXLColumn[114] = 36;   // 이연퇴직소득세(농어촌특별세) 
            
            pXLColumn[115] = 13;   // 차감원청징수세액(소득세)                                   
            pXLColumn[116] = 21;   // 차감원청징수세액(지방소득세)                   
            pXLColumn[117] = 28;   // 차감원청징수세액(농어촌특별세)  
            pXLColumn[118] = 36;   // 차감원청징수세액(농어촌특별세) 
            
            pXLColumn[119] = 35;   // 출력일자                       
            pXLColumn[120] = 27;   // 대표이사                                    
            //pXLColumn[121] = 12;   // 정산지급일

            pXLColumn[122] = 24;   // 31.퇴직소득(정산(합산))
            pXLColumn[123] = 24;   // 32.근속연수공제
            pXLColumn[124] = 24;   // 33.환산급여
            pXLColumn[125] = 24;   // 34.환산급여별공제
            pXLColumn[126] = 24;   // 35.퇴직소득과세표준
            pXLColumn[127] = 24;   // 42.환산산출세액
            pXLColumn[128] = 24;   // 43.산출세액
            pXLColumn[129] = 24;   // 44.퇴직일이 속하는 과세연도
            pXLColumn[130] = 24;   // 45.특례적용산출세액
            pXLColumn[131] = 24;   // 46.기납부 세액
            pXLColumn[132] = 24;   // 47.신고대상세액
            pXLColumn[133] = 29;   // 연금계좌 입금 합계 
            pXLColumn[134] = 2;    // 세무서장 또는 성명
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

        #region ----- Line Write Method -----

        private int XLLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, object pPrintType, string pCourse)
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 거주 구분(거주자1)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 거주 구분(거주자2)
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 내외국인 구분(내국인1) 
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

                // 내외국인 구분(외국인9) 
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 거주지국.
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 거주지코드
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 출력 용도 구분
                vXLColumnIndex = 15;
                vObject = pPrintType;
                IsConvert = IsConvertString(vObject, out vConvertString);

                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                // 징수의무지구분.
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


     //---------------------------------------------------------------------------------------------------//
                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
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

                // 법인명(상호)   
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

                // 대표자(성명)   
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 소재지(주소) 
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

                // 성명
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

                // 주민번호
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

                // 임원여부
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 주소
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 귀속연도 시작 일자      
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

                // 정년퇴직      
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

                // 정리해고     
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

                // 자발적퇴직 
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

                // 귀속연도 마지막 일자(퇴직일자)       
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

                // 임원퇴직     
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

                // 중간정산     
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

                // 기타     
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
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                // 근무처명-주(현)
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

                // 근무처명- 종(전)
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

                // 근무처명- 종(전)
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

                // 사업자등록번호1
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

                // 사업자등록번호2
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

                // 사업자등록번호3
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // 퇴직급여1-법정
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여2-법정
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여3-법정
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여 합계
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여1(명퇴수당 등)-법정외 
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여2(명퇴수당 등)-법정외 
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여3(명퇴수당 등)-법정외 
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직연금일시금1
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금일시금2
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금일시금3
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금일시금 합계
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 명예퇴직수당 퇴직연금일시금 1
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 명예퇴직수당 퇴직연금일시금2
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 명예퇴직수당 퇴직연금일시금3
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 명예퇴직수당 퇴직연금일시금 합계  
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 법정계1
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법정계2
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법정계3
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법정계 합계
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법정외계1
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법정외계2
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법정외계3
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 비과세소득1
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 비과세소득2
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 비과세소득3
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 비과세소득 합계
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                //[ 법정 퇴직급여(주(현)근무지) ]   

                // 입사일1 - 법정      
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

                // 퇴사일1 - 법정
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

                // 근속월수1 - 법정
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수1 - 법정
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[86];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 가산월수1 - 법정
                vGDColumnIndex = pGDColumn[148];
                vXLColumnIndex = pXLColumn[148];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속연수1 - 법정
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[87];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // [ 법정 외 퇴직급여(주(현)근무지) ]
                // 입사일1 - 법정 외   
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

                // 퇴사일1 - 법정 외
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

                // 근속월수1 - 법정 외
                vGDColumnIndex = pGDColumn[90];
                vXLColumnIndex = pXLColumn[90];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수1 - 법정 외
                vGDColumnIndex = pGDColumn[91];
                vXLColumnIndex = pXLColumn[91];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 가산월수1 - 법정
                vGDColumnIndex = pGDColumn[149];
                vXLColumnIndex = pXLColumn[149];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속연수1 - 법정
                vGDColumnIndex = pGDColumn[92];
                vXLColumnIndex = pXLColumn[92];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

                // [ 법정 퇴직급여(종(전)근무지) ]

                // 입사일1 - 법정      
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

                // 퇴사일1 - 법정
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

                // 근속월수1 - 법정
                vGDColumnIndex = pGDColumn[95];
                vXLColumnIndex = pXLColumn[95];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수1 - 법정
                vGDColumnIndex = pGDColumn[96];
                vXLColumnIndex = pXLColumn[96];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // [ 법정 외 퇴직급여(주(현)근무지) ]
                // 입사일1 - 법정 외   
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

                // 퇴사일1 - 법정 외
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

                // 근속월수1 - 법정 외
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수1 - 법정 외
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

                // 퇴직연금계좌번호 - 주현
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

                // 퇴직연금 일시금 총수령액 - 주현
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금 원리금 합계액 - 주현
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금 소득자 불입액 - 주현
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금 소득공제액 - 주현
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금 일시금 - 주현
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직연금 일시금 총수령액 - 주현
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금 원리금 합계액 - 주현
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금 소득자 불입액 - 주현
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금 소득공제액 - 주현
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직연금 일시금 - 주현
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직연금 일시금 지급 예상액 - 법정 - 주현
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 과세이연금액(퇴직연금일시금제외) - 법정 - 주현
                vGDColumnIndex = pGDColumn[151];
                vXLColumnIndex = pXLColumn[151];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기수령한 퇴직급여액 급여액- 법정 - 주현
                vGDColumnIndex = pGDColumn[152];
                vXLColumnIndex = pXLColumn[152];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 총퇴직연금 일시금- 법정 - 주현
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 수령가능퇴직급여액 법정 - 주현
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 환산퇴직소득공제 법정 - 주현
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 환산퇴직소득과세표준 법정 - 주현
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 환산연평균과세표준 법정 - 주현
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 환산연평균산출세액 법정 - 주현
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직연금 일시금 지급 예상액 - 법정 - 종전
                vGDColumnIndex = pGDColumn[74];
                vXLColumnIndex = pXLColumn[74];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직연금 일시금 지급 예상액 - 법정외 - 주현
                vGDColumnIndex = pGDColumn[75];
                vXLColumnIndex = pXLColumn[75];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 총퇴직연금 일시금- 법정외 - 주현
                vGDColumnIndex = pGDColumn[76];
                vXLColumnIndex = pXLColumn[76];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 수령가능퇴직급여액 법정외 - 주현
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 환산퇴직소득공제 법정외 - 주현
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 환산퇴직소득과세표준 법정외 - 주현
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 환산연평균과세표준 법정외 - 주현
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 환산연평균산출세액 법정외 - 주현
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직연금 일시금 지급 예상액 - 법정외 - 종전
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직연금사업자명
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

                // 퇴직연금사업자등록번호
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

                // 퇴직연금계좌번호
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

                // 법정퇴직급여금액
                vGDColumnIndex = pGDColumn[156];
                vXLColumnIndex = pXLColumn[156];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법정외퇴직급여금액
                vGDColumnIndex = pGDColumn[157];
                vXLColumnIndex = pXLColumn[157];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 입금일
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

                // 만기일
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

                // 과세이연세액
                vGDColumnIndex = pGDColumn[160];
                vXLColumnIndex = pXLColumn[160];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직금여액 - 법정
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직금여액 - 법정 외
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직금여액 - 계
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직소득공제 - 법정
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득공제 - 법정 외
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득공제 - 계
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 퇴직소득과세표준 - 법정
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득과세표준 - 법정 외
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직소득과세표준 - 계
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 연평균과세표준 - 법정
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 연평균과세표준 - 법정 외
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 연평균과세표준 - 계
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 연평균산출세액 - 법정
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 연평균산출세액 - 법정 외
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 연평균산출세액 - 계
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 산출세액 - 법정
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 산출세액 - 법정 외
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 산출세액 - 계
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 산출세액 - 법정
                vGDColumnIndex = pGDColumn[121];
                vXLColumnIndex = pXLColumn[121];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 산출세액 - 법정 외
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 산출세액 - 계
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 결정세액 - 소득세
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 결정세액 - 지방소득세
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 결정세액 - 농어촌특별세
                vGDColumnIndex = pGDColumn[129];
                vXLColumnIndex = pXLColumn[129];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 결정세액 - 계
                vGDColumnIndex = pGDColumn[130];
                vXLColumnIndex = pXLColumn[130];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 종전근무지기납부세액 - 소득세
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 종전근무지기납부세액 - 지방소득세
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 종전근무지기납부세액 - 농어촌특별세
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 종전근무지기납부세액 - 계
                vGDColumnIndex = pGDColumn[134];
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 차감원천징수세액 - 소득세
                vGDColumnIndex = pGDColumn[135];
                vXLColumnIndex = pXLColumn[135];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원천징수세액 - 지방소득세
                vGDColumnIndex = pGDColumn[136];
                vXLColumnIndex = pXLColumn[136];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원천징수세액 - 농어촌특별세
                vGDColumnIndex = pGDColumn[137];
                vXLColumnIndex = pXLColumn[137];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원천징수세액 - 계
                vGDColumnIndex = pGDColumn[138];
                vXLColumnIndex = pXLColumn[138];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                
                // 출력날짜
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

                // 출력날짜
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

        #region ----- Line Write Method -----
        private int XLLine2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_2013, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pPrintType, string pPrintType_Desc, string pCourse)
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
                mPrinting.XLActiveSheet("Destination");

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 거주 구분(거주자1)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 거주 구분(거주자2)
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 내외국인 구분(내국인1) 
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 내외국인 구분(외국인9) 
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 거주지국.
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 거주지코드
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 출력 용도 구분
                vXLColumnIndex = 15;
                vObject = pPrintType_Desc;
                IsConvert = IsConvertString(vObject, out vConvertString);

                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 징수의무지구분.
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                //---------------------------------------------------------------------------------------------------//
                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // 사업자등록번호 
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법인명(상호)   
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 법인번호
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 소재지 주소
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 임원여부 
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 확정급여형 퇴직연금 제도 가입일  
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직금 날짜 
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
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

                // 귀속연도 시작 일자 
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 정년퇴직.
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 정리해고
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 자발적퇴직
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 귀속연도 마지막 일자(퇴직일자). 
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 임원퇴직
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //중간정산
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //기타
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                //----퇴직급여현황----

                // 근무처명 ( 중간지급)
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무처명 (최종분)  
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 사업자등록번호 ( 중간지급)  
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 사업자등록번호 (최종분)          
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 퇴직급여 (중간지급)
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여 (최종분)  
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여 (정산)
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 비과세 퇴직급여 (중간지급)       
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 비과세 퇴직급여 (최종분)    
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 비과세 퇴직급여 (정산)   
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 과세대상 퇴직급여 (중간지급)
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 과세대상 퇴직급 (최종분)   
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 과세대상 퇴직급 (정산) 
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                //-------근속연수------------------
                // 입사일 ( 중간지급 근속연수) 
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기산일/정산시작일 ( 중간지급 근속연수)  
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴사일 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 지급일 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // 제외월수 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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


                // 가감월수 ( 중간지급 근속연수)  
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // 근속연수 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 입사일(최종분)       
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기산일/정산시작일 ( 최종분)   
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴사일 ( 최종분 )  
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 지급일 ( 최종분 )
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수 ( 최종분)
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수 ( 최종분) 
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 가산월수 ( 최종분)    
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속연수 ( 최종분)  
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

                // 입사일(정산(합산))  
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기산일/정산시작일 (정산(합산))   
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴사일 (정산(합산))
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수 (정산(합산))
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수 (정산(합산))  
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 가감월수 (정산(합산))
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 중복월수 (정산(합산)) 
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속연수 (정산(합산)) 
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

                // 기산일 (2012.12.31 이전)  
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴사일  (2012.12.31 이전) 
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수 (2012.12.31 이전)   
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                // 제외월수 (2012.12.31 이전)
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 가감월수 (2012.12.31 이전) 
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속년수 (2012.12.31 이전)
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

                // 기산일 ( 2013.01.01 이후)    
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴사일  ( 2013.01.01 이후)
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수 ( 2013.01.01 이후) 
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수 ( 2013.01.01 이후)   
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 가감월수 ( 2013.01.01 이후) 
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //근속년수 ( 2013.01.01 이후)
                vGDColumnIndex = pGDColumn[74];
                vXLColumnIndex = pXLColumn[74];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

                //----------------------- 개정규정에 따른 계산방법 ----------------------------

                // 27.퇴직소득(17)
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 28.근속연수공제
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 29.환산급여(27-28) * 12배 / 정산 근속연수   
                vGDColumnIndex = pGDColumn[124];
                vXLColumnIndex = pXLColumn[124];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 30.환산급여별공제      
                vGDColumnIndex = pGDColumn[125];
                vXLColumnIndex = pXLColumn[125];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 31.퇴직소득과세표준
                vGDColumnIndex = pGDColumn[126];
                vXLColumnIndex = pXLColumn[126];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 32.환산산출세액(31 * 세율)
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 33.산출세액(31*정산근속연수 / 12배)
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 34.퇴직소득(17)
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 35.퇴직소득정률공제
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 36.근속연수공제
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 37.퇴직소득과세표준(34-35-36)
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 38.과세표준안분(37*각근속연수/정산근속연수) - 2012.12.31 이전   
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 38.과세표준안분(37*각근속연수/정산근속연수) - 2013.01.01 이후   
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 38.과세표준안분(37*각근속연수/정산근속연수) - 합계  
                vGDColumnIndex = pGDColumn[91];
                vXLColumnIndex = pXLColumn[91];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 39.연평균과세표준(38/각근속연수) - 2012.12.31이전 
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 39.연평균과세표준(38/각근속연수) - 2013.01.01이후 
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[86];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 39.연평균과세표준(38/각근속연수) - 합계
                vGDColumnIndex = pGDColumn[92];
                vXLColumnIndex = pXLColumn[92];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertDecimal = iString.ISDecimaltoZero(vObject, 0);
                if (vConvertDecimal != 0)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //40.환산과세표준(39*5배)-2013.01.01이후 
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[87];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //40.환산과세표준(39*5배)-합계 
                vGDColumnIndex = pGDColumn[93];
                vXLColumnIndex = pXLColumn[93];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 41.환산산출세액(40*세율)-2013.01.01이후
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[88];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 41.환산산출세액(40*세율)-합계
                vGDColumnIndex = pGDColumn[94];
                vXLColumnIndex = pXLColumn[94];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 42.연평균산출세액-2012.12.31이전 
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 42.연평균산출세액-2013.01.01 이후  
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[89];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 42.연평균산출세액-합계
                vGDColumnIndex = pGDColumn[95];
                vXLColumnIndex = pXLColumn[95];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 43.산출세액(42*각근속연수)-2012.12.31 이전 
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 43.산출세액(42*각근속연수)-2013.01.01 이후  
                vGDColumnIndex = pGDColumn[90];
                vXLColumnIndex = pXLColumn[90];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 43.산출세액(42*각근속연수)-합계 
                vGDColumnIndex = pGDColumn[96];
                vXLColumnIndex = pXLColumn[96];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 44.퇴직일이 속하는 과세연도 
                vGDColumnIndex = pGDColumn[129];
                vXLColumnIndex = pXLColumn[129];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                //45.퇴직소득세 산출세액 
                vGDColumnIndex = pGDColumn[130];
                vXLColumnIndex = pXLColumn[130];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //46.기납부(또는 기과세이연) 세액 
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //47.신고대상세액(45-46)
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //48.신고대상 세액 
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 연금계산취급자
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[101];
                vXLColumnIndex = pXLColumn[101];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[102];
                vXLColumnIndex = pXLColumn[102];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 입금일  
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 계좌입금금액 
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여   
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 이연퇴직소득세   
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 52.연금계좌 입금금액 합계  
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 53.신고대상세액(소득세)  
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 신고대상세액(지방소득세)
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 신고대상세액(농어촌특별세)
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 신고대상세액(계)
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 이연퇴직소득세(소득세)  
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 이연퇴직소득세(지방소득세)  
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 이연퇴직소득세(농어촌특별세)    
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 이연퇴직소득세(계)
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 차감원청징수세액(소득세)
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원청징수세액(지방소득세)  
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원청징수세액(농어촌특별세)  
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원청징수세액(계)  
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 출력일자
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 대표이사    
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                // 세무서 단,소득자 보관용이면 소득자명 인쇄
                if (pPrintType == "1")
                {
                    vGDColumnIndex = pGDColumn[12];
                }
                else
                {
                    vGDColumnIndex = pGDColumn[134];
                }
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            pXLine = vXLine;

            return pXLine;
        }

        private int XLLine3(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_2013, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pPrintType, string pPrintType_Desc, string pCourse)
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
                mPrinting.XLActiveSheet("Destination");

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // 거주 구분(거주자1)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 거주 구분(거주자2)
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 내외국인 구분(내국인1) 
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 내외국인 구분(외국인9) 
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                ///-------종교관련종사자여부
               
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // 거주지국.
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 거주지코드
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 출력 용도 구분
                vXLColumnIndex = 15;
                vObject = pPrintType_Desc;
                IsConvert = IsConvertString(vObject, out vConvertString);

                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 징수의무지구분.
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                //---------------------------------------------------------------------------------------------------//
                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // 사업자등록번호 
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 법인명(상호)   
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 법인번호
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString); 
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 소재지 주소
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 임원여부 
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 확정급여형 퇴직연금 제도 가입일  
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직금 날짜 
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
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

                // 귀속연도 시작 일자 
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 정년퇴직.
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 정리해고
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 자발적퇴직
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 귀속연도 마지막 일자(퇴직일자). 
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 임원퇴직
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //중간정산
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //기타
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                //----퇴직급여현황----

                // 근무처명 ( 중간지급)
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근무처명 (최종분)  
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 사업자등록번호 ( 중간지급)  
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 사업자등록번호 (최종분)          
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 퇴직급여 (중간지급)
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여 (최종분)  
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여 (정산)
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 비과세 퇴직급여 (중간지급)       
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 비과세 퇴직급여 (최종분)    
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 비과세 퇴직급여 (정산)   
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 과세대상 퇴직급여 (중간지급)
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 과세대상 퇴직급 (최종분)   
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 과세대상 퇴직급 (정산) 
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                //-------근속연수------------------
                // 입사일 ( 중간지급 근속연수) 
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기산일/정산시작일 ( 중간지급 근속연수)  
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴사일 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 지급일 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // 제외월수 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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


                // 가감월수 ( 중간지급 근속연수)  
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // 근속연수 ( 중간지급 근속연수)
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 입사일(최종분)       
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기산일/정산시작일 ( 최종분)   
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴사일 ( 최종분 )  
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 지급일 ( 최종분 )
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수 ( 최종분)
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수 ( 최종분) 
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 가산월수 ( 최종분)    
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속연수 ( 최종분)  
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

                // 입사일(정산(합산))  
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기산일/정산시작일 (정산(합산))   
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴사일 (정산(합산))
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속월수 (정산(합산))
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 제외월수 (정산(합산))  
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 가감월수 (정산(합산))
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 중복월수 (정산(합산)) 
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 근속연수 (정산(합산)) 
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

               

                //----------------------- 개정규정에 따른 계산방법 ----------------------------
                
                // 27.퇴직소득(17)
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 28.근속연수공제
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 29.환산급여(27-28) * 12배 / 정산 근속연수   
                vGDColumnIndex = pGDColumn[124];
                vXLColumnIndex = pXLColumn[124];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 30.환산급여별공제      
                vGDColumnIndex = pGDColumn[125];
                vXLColumnIndex = pXLColumn[125];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 31.퇴직소득과세표준
                vGDColumnIndex = pGDColumn[126];
                vXLColumnIndex = pXLColumn[126];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 32.환산산출세액(31 * 세율)
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 33.산출세액(31*정산근속연수 / 12배)
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //// 34.퇴직소득(17)
                //vGDColumnIndex = pGDColumn[77];
                //vXLColumnIndex = pXLColumn[77];
                //vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                ////-------------------------------------------------------------------
                //vXLine = vXLine + 2;
                ////-------------------------------------------------------------------
                //// 35.퇴직소득정률공제
                //vGDColumnIndex = pGDColumn[78];
                //vXLColumnIndex = pXLColumn[78];
                //vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                ////45.퇴직소득세 산출세액 
                //vGDColumnIndex = pGDColumn[130];
                //vXLColumnIndex = pXLColumn[130];
                //vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                ////45.퇴직소득세 산출세액 
                //vGDColumnIndex = pGDColumn[130];
                //vXLColumnIndex = pXLColumn[130];
                //vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //46.기납부(또는 기과세이연) 세액 
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //47.신고대상세액(45-46)
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //48.신고대상 세액 
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 연금계산취급자
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[101];
                vXLColumnIndex = pXLColumn[101];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[102];
                vXLColumnIndex = pXLColumn[102];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 입금일  
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 계좌입금금액 
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 퇴직급여   
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 이연퇴직소득세   
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 52.연금계좌 입금금액 합계  
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 53.신고대상세액(소득세)  
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 신고대상세액(지방소득세)
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 신고대상세액(농어촌특별세)
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 신고대상세액(계)
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 이연퇴직소득세(소득세)  
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 이연퇴직소득세(지방소득세)  
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 이연퇴직소득세(농어촌특별세)    
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 이연퇴직소득세(계)
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 차감원청징수세액(소득세)
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원청징수세액(지방소득세)  
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원청징수세액(농어촌특별세)  
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 차감원청징수세액(계)  
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // 출력일자
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 대표이사    
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                // 세무서 단,소득자 보관용이면 소득자명 인쇄
                if (pPrintType == "1")
                {
                    vGDColumnIndex = pGDColumn[12];
                }
                else
                {
                    vGDColumnIndex = pGDColumn[134];
                }
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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


        #region ----- Excel Write WithholdingTax  Method ----

        public int WriteWithholdingTax(string pPrint_Type, string pSaveFileName, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_2013, string pPrintType, string pPrintType_Desc)
        {
            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn;
            int[] vXLColumn;

            int vTotalRow = pGrid_WITHHOLDING_TAX.RowCount;
            int vTotalRow2  = pGrid_PRINT_2013.RowCount;
            string vRetire_Year = pGrid_PRINT_2013.GetCellValue("FINAL_RETIRE_DATE").ToString();


            int vRowCount = 0;

            int vPrintingLine = 0;

            int vSecondPrinting = 9;
            int vCountPrinting = 0;

            if (iString.ISNumtoZero(vRetire_Year) < 2013)
            {
                SetArray1(pGrid_WITHHOLDING_TAX, out vGDColumn, out vXLColumn);

                for (int vRow = 0; vRow < vTotalRow; vRow++)
                {
                    vRowCount++;
                    pGrid_WITHHOLDING_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                    mAppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();

                    if (isOpen == true)
                    {
                        vCountPrinting++;

                        mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                        pGrid_WITHHOLDING_TAX.CurrentCellMoveTo(vRow, 0);
                        pGrid_WITHHOLDING_TAX.Focus();
                        pGrid_WITHHOLDING_TAX.CurrentCellActivate(vRow, 0);

                        // 퇴직소득원천징수영수증/지급조서
                        vPrintingLine = XLLine(pGrid_WITHHOLDING_TAX, vRow, vPrintingLine, vGDColumn, vXLColumn, pPrintType, "SRC_TAB1");

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
            }
            else if(iString.ISNumtoZero(vRetire_Year) < 2020)
            {
                SetArray2(pGrid_PRINT_2013, out vGDColumn, out vXLColumn);

                for (int vRow = 0; vRow < vTotalRow; vRow++)
                {
                    vRowCount++;
                    pGrid_WITHHOLDING_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                    mAppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();

                    if (isOpen == true)
                    {
                        vCountPrinting++;

                        mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                        pGrid_PRINT_2013.CurrentCellMoveTo(vRow, 0);
                        pGrid_PRINT_2013.Focus();
                        pGrid_PRINT_2013.CurrentCellActivate(vRow, 0);

                        // 퇴직소득원천징수영수증/지급조서
                        vPrintingLine = XLLine2(pGrid_PRINT_2013, vRow, vPrintingLine, vGDColumn, vXLColumn, pPrintType, pPrintType_Desc, "SRC_TAB1");

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
            }
            else
            {
                SetArray2(pGrid_PRINT_2013, out vGDColumn, out vXLColumn);

                for (int vRow = 0; vRow < vTotalRow; vRow++)
                {
                    vRowCount++;
                    pGrid_WITHHOLDING_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                    mAppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();

                    if (isOpen == true)
                    {
                        vCountPrinting++;

                        mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                        pGrid_PRINT_2013.CurrentCellMoveTo(vRow, 0);
                        pGrid_PRINT_2013.Focus();
                        pGrid_PRINT_2013.CurrentCellActivate(vRow, 0);

                        // 퇴직소득원천징수영수증/지급조서
                        vPrintingLine = XLLine3(pGrid_PRINT_2013, vRow, vPrintingLine, vGDColumn, vXLColumn, pPrintType, pPrintType_Desc, "SRC_TAB1");

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
            }
            //SAVE("RETIRE.XLS");
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

            if (pCourse == "SRC_TAB1")
            {
                pPrinting.XLActiveSheet("SourceTab1");
            }            

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