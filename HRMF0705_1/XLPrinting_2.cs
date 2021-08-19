using System;
using System.Collections.Generic;
using System.Text;

namespace HRMF0705
{
    class XLPrinting_2
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

        private int mPrintingLineSTART = 6;  //Line

        private int mCopyLineSUM = 1;        //엑셀의 선택된 쉬트의 복사되어질 시작 행 위치, 복사 행 누적
        private int mIncrementCopyMAX = 67;  // 1page : 61, 2page : 122, 3page : 183 - 복사되어질 행의 범위

        private int mCopyColumnSTART = 1;    //복사되어  진 행 누적 수
        private int mCopyColumnEND = 45;     //엑셀의 선택된 쉬트의 복사되어질 끝 열 위치

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

        public XLPrinting_2(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
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

        #region ----- Array Set 3 ----
        private void SetArray3(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_INCOME_TAX, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[60];
            pXLColumn = new int[60];

            pGDColumn[0] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("NAME");  // 다자녀 인원수
            pGDColumn[1] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("REPRE_NUM");         // 관계코드         
            pGDColumn[2] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("ADDRESS");           // 성명             
            pGDColumn[3] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("CORP_NAME");               // 기본공제         
            pGDColumn[4] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("VAT_NUMBER");                // 경로우대         
            pGDColumn[5] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("CORP_ADDRESS");              // 출산/입양양육    
            pGDColumn[6] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TEL_NUMBER");             // 장애인           
            pGDColumn[7] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PRESIDENT_NAME");              // 자녀양육(6세이하)
            pGDColumn[8] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("LEGAL_NUMBER");
            // 국세청-보험료    
            pGDColumn[9] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_01");           // 국세청-의료비    
            pGDColumn[10] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_01");               // 국세청-교육비    
            pGDColumn[11] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_01");            // 국세청-신용카드  
            pGDColumn[12] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_01");

            pGDColumn[13] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_08");           // 국세청-의료비    
            pGDColumn[14] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_08");               // 국세청-교육비    
            pGDColumn[15] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_08");            // 국세청-신용카드  
            pGDColumn[16] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_08");

            pGDColumn[17] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_02");           // 국세청-의료비    
            pGDColumn[18] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_02");               // 국세청-교육비    
            pGDColumn[19] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_02");            // 국세청-신용카드  
            pGDColumn[20] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_02");

            pGDColumn[21] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_09");           // 국세청-의료비    
            pGDColumn[22] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_09");               // 국세청-교육비    
            pGDColumn[23] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_09");            // 국세청-신용카드  
            pGDColumn[24] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_09");

            pGDColumn[25] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_03");           // 국세청-의료비    
            pGDColumn[26] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_03");               // 국세청-교육비    
            pGDColumn[27] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_03");            // 국세청-신용카드  
            pGDColumn[28] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_03");

            pGDColumn[29] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_10");           // 국세청-의료비    
            pGDColumn[30] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_10");               // 국세청-교육비    
            pGDColumn[31] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_10");            // 국세청-신용카드  
            pGDColumn[32] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_10");

            pGDColumn[33] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_04");           // 국세청-의료비    
            pGDColumn[34] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_04");               // 국세청-교육비    
            pGDColumn[35] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_04");            // 국세청-신용카드  
            pGDColumn[36] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_04");

            pGDColumn[37] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_11");           // 국세청-의료비    
            pGDColumn[38] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_11");               // 국세청-교육비    
            pGDColumn[39] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_11");            // 국세청-신용카드  
            pGDColumn[40] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_11");

            pGDColumn[41] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_05");           // 국세청-의료비    
            pGDColumn[42] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_05");               // 국세청-교육비    
            pGDColumn[43] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_05");            // 국세청-신용카드  
            pGDColumn[44] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_05");

            pGDColumn[45] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_12");           // 국세청-의료비    
            pGDColumn[46] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_12");               // 국세청-교육비    
            pGDColumn[47] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_12");            // 국세청-신용카드  
            pGDColumn[48] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_12");

            pGDColumn[49] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_06");           // 국세청-의료비    
            pGDColumn[50] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_06");               // 국세청-교육비    
            pGDColumn[51] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_06");            // 국세청-신용카드  
            pGDColumn[52] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_06");

            pGDColumn[53] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("PAY_YYYYMM_07");           // 국세청-의료비    
            pGDColumn[54] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_AMOUNT_07");               // 국세청-교육비    
            pGDColumn[55] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TAX_AMOUNT_07");            // 국세청-신용카드  
            pGDColumn[56] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("SUPPLY_DATE_07");

            pGDColumn[57] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TOTAL_SUPPLY_AMOUNT");              // 국세청-현금      
            pGDColumn[58] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("TOTAL_TAX_AMOUNT");             // 국세청-기부금    
            pGDColumn[59] = pGrid_PRINT_INCOME_TAX.GetColumnToIndex("NAME");             // 국세청-기부금   

            pXLColumn[0] = 13;
            pXLColumn[1] = 35;
            pXLColumn[2] = 13;
            pXLColumn[3] = 13;
            pXLColumn[4] = 35;
            pXLColumn[5] = 13;
            pXLColumn[6] = 36;
            pXLColumn[7] = 13;
            pXLColumn[8] = 35;



            pXLColumn[9] = 2;
            pXLColumn[10] = 7;
            pXLColumn[11] = 13;
            pXLColumn[12] = 18;

            pXLColumn[13] = 24;
            pXLColumn[14] = 29;
            pXLColumn[15] = 35;
            pXLColumn[16] = 40;

            pXLColumn[17] = 2;
            pXLColumn[18] = 7;
            pXLColumn[19] = 13;
            pXLColumn[20] = 18;

            pXLColumn[21] = 24;
            pXLColumn[22] = 29;
            pXLColumn[23] = 35;
            pXLColumn[24] = 40;

            pXLColumn[25] = 2;
            pXLColumn[26] = 7;
            pXLColumn[27] = 13;
            pXLColumn[28] = 18;

            pXLColumn[29] = 24;
            pXLColumn[30] = 29;
            pXLColumn[31] = 35;
            pXLColumn[32] = 40;

            pXLColumn[33] = 2;
            pXLColumn[34] = 7;
            pXLColumn[35] = 13;
            pXLColumn[36] = 18;

            pXLColumn[37] = 24;
            pXLColumn[38] = 29;
            pXLColumn[39] = 35;
            pXLColumn[40] = 40;

            pXLColumn[41] = 2;
            pXLColumn[42] = 7;
            pXLColumn[43] = 13;
            pXLColumn[44] = 18;

            pXLColumn[45] = 24;
            pXLColumn[46] = 29;
            pXLColumn[47] = 35;
            pXLColumn[48] = 40;

            pXLColumn[49] = 2;
            pXLColumn[50] = 7;
            pXLColumn[51] = 13;
            pXLColumn[52] = 18;

            pXLColumn[53] = 2;
            pXLColumn[54] = 7;
            pXLColumn[55] = 13;
            pXLColumn[56] = 18;

            pXLColumn[57] = 29;
            pXLColumn[58] = 35;
            pXLColumn[59] = 21;

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
            mPrinting.XLActiveSheet("SourceTab3");
            int vXLine = 21;
            int vXLColumnIndex = 29;
            mPrinting.XLSetCell(vXLine, vXLColumnIndex, mSend_ORG);

            vXLColumnIndex = 41;
            mPrinting.XLSetCell(vXLine, vXLColumnIndex, mPrint_COUNT);
        }

        #endregion;

        #region -----XLLINE3 -----

        private int XLLine3(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_INCOME_TAX, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pCourse)
        {
            int vXLine = pXLine; // 엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;

            //decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");




                vXLine = vXLine + 0;

                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                // 관계코드
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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


                // 성명
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 기본공제
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                //vXLine = 63;
                //vXLColumnIndex = 24;
                //mPrinting.XLSetCell(63, 24, vObject);


                // 경로우대
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 출산/입양양육
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                // 장애인
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 자녀양육(6세이하)
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 국세청-의료비
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 국세청-교육비
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 국세청-신용카드
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 국세청-직불카드
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 국세청-현금
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                // 국가타입
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                // 부녀자
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기타-의료비
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기타-교육비
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기타-신용카드
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기타-직불카드
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 기타-현금
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
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
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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

                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                vGDColumnIndex = pGDColumn[3]; 
                vObject = pGrid_PRINT_INCOME_TAX.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, 24, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, 24, vConvertString);
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

        #region ----- Excel Main Wirte  Method Backup----

        //public int WriteMain(InfoSummit.Win.ControlAdv.ISGridAdvEx pgridPRINT_INCOME_TAX, object vPrintDate, object vPrintType, string pSend_ORG, string pPrint_COUNT, string pDate)
        //{
        //    string vMessageText = string.Empty;
        //    bool isOpen = XLFileOpen();
        //    mCopyLineSUM = 1;
        //    mPageNumber = 0;

        //    mSend_ORG = pSend_ORG;
        //    mPrint_COUNT = pPrint_COUNT;

        //    int[] vGDColumn_1;
        //    int[] vXLColumn_1;



        //    int vTotalRow3 = pgridPRINT_INCOME_TAX.RowCount;
        //    int vRowCount = 0;

        //    int vPrintingLine = 0;

        //    int vSecondPrinting = 30; //1인당 3페이지이므로, 3*10=30번째에 인쇄
        //    int vCountPrinting = 0;

        //    XLHeader(pDate);


        //    SetArray3(pgridPRINT_INCOME_TAX, out vGDColumn_1, out vXLColumn_1);
        //    SendORG();

        //        for (int vRow1 = 0; vRow1 < vTotalRow3; vRow1++)
        //        {
        //            vRowCount++;
        //            pgridPRINT_INCOME_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

        //            vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow3);
        //            mAppInterface.OnAppMessageEvent(vMessageText);
        //            System.Windows.Forms.Application.DoEvents();

        //            if (isOpen == true)
        //            {
        //                vCountPrinting++;

        //                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
        //                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

        //                pgridPRINT_INCOME_TAX.CurrentCellMoveTo(vRow1, 0);
        //                pgridPRINT_INCOME_TAX.Focus();
        //                pgridPRINT_INCOME_TAX.CurrentCellActivate(vRow1, 0);

        //                // 근로소득원천징수영수증 page 1 - 2.
        //                //int vLinePrinting_1 = vPrintingLine + 3;


        //                vPrintingLine = XLLine3(pgridPRINT_INCOME_TAX, vRow1, vPrintingLine, vGDColumn_1, vXLColumn_1, "SRC_TAB1");

        //                //// 부양가족내역 page 3.
        //                //int vPrintingLine_2 = vPrintingLine + 8;
        //                //for (int vRow2 = 0; vRow2 < vTotalRow2; vRow2++)
        //                //{
        //                //    vPrintingLine = XLLine2(pGrid_SUPPORT_FAMILY, vRow2, vPrintingLine, vGDColumn_2, vXLColumn_2, "SRC_TAB1");
        //                //}

        //                if (vSecondPrinting < vCountPrinting)
        //                {
        //                    Printing(1, vSecondPrinting);

        //                    mPrinting.XLOpenFileClose();
        //                    isOpen = XLFileOpen();

        //                    vCountPrinting = 0;
        //                    vPrintingLine = 1;
        //                    mCopyLineSUM = 1;
        //                }
        //                else if (vTotalRow3 == vRowCount)
        //                {
        //                    Printing(1, (vCountPrinting * 3)); //vSecondPrinting, (vRowCount * 3) 1인은 1쪽부터 해당 쪽수까지 한페이지로 보기 때문에 *2를 한 것임.
        //                }
        //            }
        //    }
        //    mPrinting.XLOpenFileClose();

        //    return mPageNumber;
        //}

        #endregion;

        #region ----- Excel Main Wirte  Method Backup----

        public int WriteMain(InfoSummit.Win.ControlAdv.ISGridAdvEx pgridPRINT_INCOME_TAX, object vPrintDate, object vPrintType, string pSend_ORG, string pPrint_COUNT, string pDate)
        {
            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            mSend_ORG = pSend_ORG;
            mPrint_COUNT = pPrint_COUNT;

            int[] vGDColumn_1;
            int[] vXLColumn_1;


            int vTotalRow3 = pgridPRINT_INCOME_TAX.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 0;

            XLHeader(pDate);

            SetArray3(pgridPRINT_INCOME_TAX, out vGDColumn_1, out vXLColumn_1);
            SendORG();

            for (int vRow1 = 0; vRow1 < vTotalRow3; vRow1++)
            {
                vRowCount++;
                pgridPRINT_INCOME_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow3);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                pgridPRINT_INCOME_TAX.CurrentCellMoveTo(vRow1, 0);
                pgridPRINT_INCOME_TAX.Focus();
                pgridPRINT_INCOME_TAX.CurrentCellActivate(vRow1, 0);

                vPrintingLine = XLLine3(pgridPRINT_INCOME_TAX, vRow1, vPrintingLine, vGDColumn_1, vXLColumn_1, "SRC_TAB1");
            }

            return mPageNumber;
        }

        #endregion;

        #endregion;

        #region ----- Header Write Method ----

        private void XLHeader(string pDate)
        {
            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet("SourceTab3");

                vXLine = 53;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, pDate);

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

            if (pCourse == "SRC_TAB1")
            {
                pPrinting.XLActiveSheet("SourceTab3");
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
