using System;
using ISCommonUtil;

namespace HRMF0502
{
    public class XL_Upload
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private string mMessageError = string.Empty;

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;
        
        public XL.XLPrint mExcel_Upload = null;

        private string mXLOpenFileName = string.Empty;

        private int mTotalROW = 0;    //Excel Active Sheet Row Count
        private int mTotalCOLUMN = 0; //Excel Active Sheet Column Count

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        public string OpenFileName
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        public int TotalROW
        {
            get
            {
                return mTotalROW;
            }
            set
            {
                mTotalROW = value;
            }
        }

        public int TotalCOLUMN
        {
            get
            {
                return mTotalCOLUMN;
            }
            set
            {
                mTotalCOLUMN = value;
            }
        }

        //public int ReadRow
        //{
        //    get
        //    {
        //        return mStartRowRead;
        //    }
        //    set
        //    {
        //        mStartRowRead = value;
        //    }
        //}

        #endregion;

        #region ----- Constructor -----

        public XL_Upload()
        {
            mExcel_Upload = new XL.XLPrint();
        }

        public XL_Upload(InfoSummit.Win.ControlAdv.ISAppInterfaceAdv pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mAppInterface = pAppInterface;
            mMessageAdapter = pMessageAdapter;

            mExcel_Upload = new XL.XLPrint();
        }

        #endregion;

        #region ----- XLDispose -----

        public void DisposeXL()
        {
            mExcel_Upload.XLOpenFileClose();
            mExcel_Upload.XLClose();
        }

        #endregion;

        #region ----- XL File Open -----

        public bool OpenXL()
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mExcel_Upload.XLFileOpen(mXLOpenFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }

            return IsOpen;
        }

        #endregion;

        #region ----- Convert String Methods ----

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
            catch
            {
            }

            return vString;
        }

        #endregion;

        #region ----- Convert Date Methods ----

        private System.DateTime ConvertDate(object pObject)
        {
            bool isConvert = false;
            string vTextDateTimeShort = string.Empty;
            System.DateTime vDate = DateTime.Today;

            try
            {
                if (pObject != null)
                {
                    isConvert = pObject is double;
                    if (isConvert == true)
                    {
                        double isConvertDouble = (double)pObject;
                        vDate = System.DateTime.FromOADate(isConvertDouble);
                    }
                    else if (iDate.ISDate(pObject) == true)
                    {
                        vDate = iDate.ISGetDate(pObject);
                    }
                    else
                    {
                        vDate = iDate.ISGetDate("-");
                    }
                }
            }
            catch
            {
                vDate = iDate.ISGetDate("-");
            }
            return vDate;
        }

        #endregion;

        #region ----- Convert Decimal Methods ----

        private decimal ConvertDecimal(object pObject)
        {
            bool isConvert = false;
            decimal vConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    isConvert = pObject is decimal;
                    if (isConvert == true)
                    {
                        decimal isConvertNum = (decimal)pObject;
                        vConvertDecimal = isConvertNum;
                    }
                }

            }
            catch
            {

            }
            return vConvertDecimal;
        }

        #endregion;

        #region ----- Convert Double Methods ----

        private decimal ConvertDouble(object pObject)
        {
            bool isConvert = false;
            decimal vConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    isConvert = pObject is double;
                    if (isConvert == true)
                    {
                        double isConvertDouble = (double)pObject;
                        vConvertDecimal = Convert.ToDecimal(isConvertDouble);
                    }
                }
            }
            catch
            {
            }

            return vConvertDecimal;
        }

        #endregion;

        #region ----- XL Loading -----

        public bool LoadXL(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pStartRow)
        {
            string vMessage = string.Empty;

            
            mExcel_Upload.XLActiveSheet(1);
            mTotalROW = mExcel_Upload.CountROW + 1;
            mTotalCOLUMN = pAdapter.SelectColElement.Count;

            bool isLoad = false;
            System.Type vType = null;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vConvertDate = new DateTime();

            object vPERSON_NUM = string.Empty;

            int vADRow = 0;
            int vADCol = 0;

            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {
                    pAdapter.AddUnder();
                    //KEY값에 해당하는 셀에 DATA가 있을 경우만 INSERT를 처리해야 하므로//
                    vType = pAdapter.CurrentRow.Table.Columns["PERSON_NUM"].DataType;
                    if (vType.Name == "String")
                    {
                        vObject = mExcel_Upload.XLGetCell(vRow, 2);  //사원번호.
                        vPERSON_NUM = iString.ISNull(vObject);
                    }
                    else 
                    {
                        vPERSON_NUM = string.Empty;
                        pAdapter.Delete();
                    }
                    if (iString.ISNull(vPERSON_NUM) != string.Empty)  //사원번호가 있을 경우만 처리.
                    {                        
                        for (int vCol = 1; vCol < mTotalCOLUMN; vCol++)
                        {
                            vType = pAdapter.CurrentRow.Table.Columns[vADCol].DataType;
                            vObject = mExcel_Upload.XLGetCell(vRow, vCol);
                            if (vType != null)
                            {
                                if (iString.ISNull(vObject) == string.Empty)
                                {
                                    pAdapter.CurrentRow[vADCol] = DBNull.Value;
                                }
                                else if (vType.Name == "String")
                                {
                                    vConvertString = iString.ISNull(vObject);
                                    vConvertString = vConvertString.Trim();
                                    pAdapter.CurrentRow[vADCol] = vConvertString;
                                }
                                else if (vType.Name == "Decimal")
                                {
                                    vConvertDecimal = iString.ISDecimaltoZero(vObject);
                                    pAdapter.CurrentRow[vADCol] = vConvertDecimal;
                                }
                                else if (vType.Name == "Double")
                                {
                                    vConvertDecimal = ConvertDouble(vObject);
                                    pAdapter.CurrentRow[vADCol] = vConvertDecimal;
                                }
                                else if (vType.Name == "DateTime")
                                {
                                    vConvertDate = ConvertDate(vObject);
                                    if(vConvertDate == iDate.ISGetDate("-"))
                                    {
                                        pAdapter.CurrentRow[vADCol] = DBNull.Value;
                                    }
                                    else
                                    {
                                        pAdapter.CurrentRow[vADCol] = vConvertDate;
                                    }
                                }
                            }
                            vADCol++;
                        }
                    }
                    vADRow++;
                    vADCol = 0;

                    vMessage = string.Format("Excel Uploading : {0:D4}/{1:D4}", vRow, (mTotalROW - 1));
                    mAppInterface.OnAppMessage(vMessage);
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();
                }
                isLoad = true;
            }
            catch (System.Exception ex)
            {
                DisposeXL();

                mAppInterface.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return isLoad;
        }

        #endregion;

        #region ----- XL Loading -----

        public bool LoadXL(InfoSummit.Win.ControlAdv.ISDataCommand pCMD, int pStartRow, InfoSummit.Win.ControlAdv.ISProgressBar pPB, InfoSummit.Win.ControlAdv.ISPrompt pPM)
        {
            string vMessage = string.Empty;
            string vErrorMessage = string.Empty;
             
            mExcel_Upload.XLActiveSheet(1);
            mTotalROW = mExcel_Upload.CountROW + 1; 

            bool isLoad = false;
              
            object vSTD_YYYYMM = string.Empty;
            object vNAME = string.Empty;
            object vPERSON_NUM = string.Empty;
            object vCHARGE_DATE = string.Empty;
            object vCHARGE_CODE = string.Empty;

            int vADRow = 0;
            int vERR_CNT = 0;
            int vItem_Start = 15;

            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {
                    //KEY값에 해당하는 셀에 DATA가 있을 경우만 INSERT를 처리해야 하므로//
                    vNAME = string.Empty;
                    vPERSON_NUM = string.Empty;
                    vSTD_YYYYMM = mExcel_Upload.XLGetCell(vRow, 1);  //근무일자.                    
                    vNAME = mExcel_Upload.XLGetCell(vRow, 2);
                    vPERSON_NUM = mExcel_Upload.XLGetCell(vRow, 3);                    
                    if (iString.ISNull(vPERSON_NUM) != string.Empty)  //사원번호가 있을 경우만 처리.
                    {
                        try
                        {                            
                            pCMD.SetCommandParamValue("P_STD_YYYYMM", vSTD_YYYYMM);
                            pCMD.SetCommandParamValue("P_NAME", vNAME); 
                            pCMD.SetCommandParamValue("P_PERSON_NUM", vPERSON_NUM); 
                            pCMD.SetCommandParamValue("P_BANK_CODE", mExcel_Upload.XLGetCell(vRow, 6));
                            pCMD.SetCommandParamValue("P_BANK_ACCOUNTS", mExcel_Upload.XLGetCell(vRow, 7));
                            pCMD.SetCommandParamValue("P_PAY_TYPE", mExcel_Upload.XLGetCell(vRow, 8));
                            pCMD.SetCommandParamValue("P_GRADE_STEP", mExcel_Upload.XLGetCell(vRow, 10));
                            pCMD.SetCommandParamValue("P_CURRENCY_CODE", mExcel_Upload.XLGetCell(vRow, 11));
                            pCMD.SetCommandParamValue("P_DED_FAMILY_COUNT", mExcel_Upload.XLGetCell(vRow, 12));
                            pCMD.SetCommandParamValue("P_INCOME_TAX_RATE", mExcel_Upload.XLGetCell(vRow, 13));
                            pCMD.SetCommandParamValue("P_A01", mExcel_Upload.XLGetCell(vRow, vItem_Start));
                            pCMD.SetCommandParamValue("P_A02", mExcel_Upload.XLGetCell(vRow, vItem_Start + 1));
                            pCMD.SetCommandParamValue("P_A03", mExcel_Upload.XLGetCell(vRow, vItem_Start + 2));
                            pCMD.SetCommandParamValue("P_A04", mExcel_Upload.XLGetCell(vRow, vItem_Start + 3));
                            pCMD.SetCommandParamValue("P_A05", mExcel_Upload.XLGetCell(vRow, vItem_Start + 4));
                            pCMD.SetCommandParamValue("P_A06", mExcel_Upload.XLGetCell(vRow, vItem_Start + 5));
                            pCMD.SetCommandParamValue("P_A07", mExcel_Upload.XLGetCell(vRow, vItem_Start + 6));
                            pCMD.SetCommandParamValue("P_A08", mExcel_Upload.XLGetCell(vRow, vItem_Start + 7));
                            pCMD.SetCommandParamValue("P_A09", mExcel_Upload.XLGetCell(vRow, vItem_Start + 8));
                            pCMD.SetCommandParamValue("P_A10", mExcel_Upload.XLGetCell(vRow, vItem_Start + 9));

                            pCMD.SetCommandParamValue("P_A11", mExcel_Upload.XLGetCell(vRow, vItem_Start + 10));
                            pCMD.SetCommandParamValue("P_A12", mExcel_Upload.XLGetCell(vRow, vItem_Start + 11));
                            pCMD.SetCommandParamValue("P_A13", mExcel_Upload.XLGetCell(vRow, vItem_Start + 12));
                            pCMD.SetCommandParamValue("P_A14", mExcel_Upload.XLGetCell(vRow, vItem_Start + 13));
                            pCMD.SetCommandParamValue("P_A15", mExcel_Upload.XLGetCell(vRow, vItem_Start + 14));
                            pCMD.SetCommandParamValue("P_A16", mExcel_Upload.XLGetCell(vRow, vItem_Start + 15));
                            pCMD.SetCommandParamValue("P_A17", mExcel_Upload.XLGetCell(vRow, vItem_Start + 16));
                            pCMD.SetCommandParamValue("P_A18", mExcel_Upload.XLGetCell(vRow, vItem_Start + 17));
                            pCMD.SetCommandParamValue("P_A19", mExcel_Upload.XLGetCell(vRow, vItem_Start + 18));
                            pCMD.SetCommandParamValue("P_A20", mExcel_Upload.XLGetCell(vRow, vItem_Start + 19));

                            pCMD.SetCommandParamValue("P_A21", mExcel_Upload.XLGetCell(vRow, vItem_Start + 20));
                            pCMD.SetCommandParamValue("P_A22", mExcel_Upload.XLGetCell(vRow, vItem_Start + 21));
                            pCMD.SetCommandParamValue("P_A23", mExcel_Upload.XLGetCell(vRow, vItem_Start + 22));
                            pCMD.SetCommandParamValue("P_A24", mExcel_Upload.XLGetCell(vRow, vItem_Start + 23));
                            pCMD.SetCommandParamValue("P_A25", mExcel_Upload.XLGetCell(vRow, vItem_Start + 24));
                            pCMD.SetCommandParamValue("P_A26", mExcel_Upload.XLGetCell(vRow, vItem_Start + 25));
                            pCMD.SetCommandParamValue("P_A27", mExcel_Upload.XLGetCell(vRow, vItem_Start + 26));
                            pCMD.SetCommandParamValue("P_A28", mExcel_Upload.XLGetCell(vRow, vItem_Start + 27));
                            pCMD.SetCommandParamValue("P_A29", mExcel_Upload.XLGetCell(vRow, vItem_Start + 28));
                            pCMD.SetCommandParamValue("P_A30", mExcel_Upload.XLGetCell(vRow, vItem_Start + 29));

                            pCMD.SetCommandParamValue("P_A31", mExcel_Upload.XLGetCell(vRow, vItem_Start + 30));
                            pCMD.SetCommandParamValue("P_A32", mExcel_Upload.XLGetCell(vRow, vItem_Start + 31));
                            pCMD.SetCommandParamValue("P_A33", mExcel_Upload.XLGetCell(vRow, vItem_Start + 32));
                            pCMD.SetCommandParamValue("P_A34", mExcel_Upload.XLGetCell(vRow, vItem_Start + 33));
                            pCMD.SetCommandParamValue("P_A35", mExcel_Upload.XLGetCell(vRow, vItem_Start + 34));
                            pCMD.SetCommandParamValue("P_A36", mExcel_Upload.XLGetCell(vRow, vItem_Start + 35));
                            pCMD.SetCommandParamValue("P_A37", mExcel_Upload.XLGetCell(vRow, vItem_Start + 36));
                            pCMD.SetCommandParamValue("P_A38", mExcel_Upload.XLGetCell(vRow, vItem_Start + 37));
                            pCMD.SetCommandParamValue("P_A39", mExcel_Upload.XLGetCell(vRow, vItem_Start + 38));
                            pCMD.SetCommandParamValue("P_A40", mExcel_Upload.XLGetCell(vRow, vItem_Start + 39));

                            pCMD.SetCommandParamValue("P_D01", mExcel_Upload.XLGetCell(vRow, vItem_Start + 41));
                            pCMD.SetCommandParamValue("P_D02", mExcel_Upload.XLGetCell(vRow, vItem_Start + 42));
                            pCMD.SetCommandParamValue("P_D03", mExcel_Upload.XLGetCell(vRow, vItem_Start + 43));
                            pCMD.SetCommandParamValue("P_D04", mExcel_Upload.XLGetCell(vRow, vItem_Start + 44));
                            pCMD.SetCommandParamValue("P_D05", mExcel_Upload.XLGetCell(vRow, vItem_Start + 45));
                            pCMD.SetCommandParamValue("P_D06", mExcel_Upload.XLGetCell(vRow, vItem_Start + 46));
                            pCMD.SetCommandParamValue("P_D07", mExcel_Upload.XLGetCell(vRow, vItem_Start + 47));
                            pCMD.SetCommandParamValue("P_D08", mExcel_Upload.XLGetCell(vRow, vItem_Start + 48));
                            pCMD.SetCommandParamValue("P_D09", mExcel_Upload.XLGetCell(vRow, vItem_Start + 49));
                            pCMD.SetCommandParamValue("P_D10", mExcel_Upload.XLGetCell(vRow, vItem_Start + 50));

                            pCMD.SetCommandParamValue("P_D11", mExcel_Upload.XLGetCell(vRow, vItem_Start + 51));
                            pCMD.SetCommandParamValue("P_D12", mExcel_Upload.XLGetCell(vRow, vItem_Start + 52));
                            pCMD.SetCommandParamValue("P_D13", mExcel_Upload.XLGetCell(vRow, vItem_Start + 53));
                            pCMD.SetCommandParamValue("P_D14", mExcel_Upload.XLGetCell(vRow, vItem_Start + 54));
                            pCMD.SetCommandParamValue("P_D15", mExcel_Upload.XLGetCell(vRow, vItem_Start + 55));
                            pCMD.SetCommandParamValue("P_D16", mExcel_Upload.XLGetCell(vRow, vItem_Start + 56));
                            pCMD.SetCommandParamValue("P_D17", mExcel_Upload.XLGetCell(vRow, vItem_Start + 57));
                            pCMD.SetCommandParamValue("P_D18", mExcel_Upload.XLGetCell(vRow, vItem_Start + 58));
                            pCMD.SetCommandParamValue("P_D19", mExcel_Upload.XLGetCell(vRow, vItem_Start + 59));
                            pCMD.SetCommandParamValue("P_D20", mExcel_Upload.XLGetCell(vRow, vItem_Start + 60));

                            pCMD.SetCommandParamValue("P_D21", mExcel_Upload.XLGetCell(vRow, vItem_Start + 61));
                            pCMD.SetCommandParamValue("P_D22", mExcel_Upload.XLGetCell(vRow, vItem_Start + 62));
                            pCMD.SetCommandParamValue("P_D23", mExcel_Upload.XLGetCell(vRow, vItem_Start + 63));
                            pCMD.SetCommandParamValue("P_D24", mExcel_Upload.XLGetCell(vRow, vItem_Start + 64));
                            pCMD.SetCommandParamValue("P_D25", mExcel_Upload.XLGetCell(vRow, vItem_Start + 65));
                            pCMD.SetCommandParamValue("P_D26", mExcel_Upload.XLGetCell(vRow, vItem_Start + 66));
                            pCMD.SetCommandParamValue("P_D27", mExcel_Upload.XLGetCell(vRow, vItem_Start + 67));
                            pCMD.SetCommandParamValue("P_D28", mExcel_Upload.XLGetCell(vRow, vItem_Start + 68));
                            pCMD.SetCommandParamValue("P_D29", mExcel_Upload.XLGetCell(vRow, vItem_Start + 69));
                            pCMD.SetCommandParamValue("P_D30", mExcel_Upload.XLGetCell(vRow, vItem_Start + 70));

                            pCMD.SetCommandParamValue("P_D31", mExcel_Upload.XLGetCell(vRow, vItem_Start + 71));
                            pCMD.SetCommandParamValue("P_D32", mExcel_Upload.XLGetCell(vRow, vItem_Start + 72));
                            pCMD.SetCommandParamValue("P_D33", mExcel_Upload.XLGetCell(vRow, vItem_Start + 73));
                            pCMD.SetCommandParamValue("P_D34", mExcel_Upload.XLGetCell(vRow, vItem_Start + 74));
                            pCMD.SetCommandParamValue("P_D35", mExcel_Upload.XLGetCell(vRow, vItem_Start + 75));
                            pCMD.ExecuteNonQuery();
                            if (iString.ISNull(pCMD.GetCommandParamValue("O_STATUS")) == "F")
                            {
                                vErrorMessage = iString.ISNull(pCMD.GetCommandParamValue("O_MESSAGE"));
                                vERR_CNT++; 
                                pPM.PromptText = string.Format("Imporing :: {0}-{1} *** {2}({3} ** Error : {4})", vADRow, mTotalROW, vNAME, vPERSON_NUM, vMessage);
                                return false;
                            }
                        }
                        catch (Exception Ex)
                        {
                            DisposeXL();

                            mAppInterface.OnAppMessage(Ex.Message);
                            System.Windows.Forms.Application.DoEvents();
                            return false;
                        }
                    }
                    vADRow++;

                    pPB.BarFillPercent = (Convert.ToSingle(vADRow) / Convert.ToSingle(mTotalROW)) * 100F;
                    pPM.PromptText = string.Format("Imporing :: {0}-{1} *** {2}({3})", vADRow, mTotalROW, vNAME, vPERSON_NUM); 

                    vMessage = string.Format("Excel Uploading : {0:D4}/{1:D4}", vRow, (mTotalROW - 1));
                    mAppInterface.OnAppMessage(vMessage);
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();
                }
                if (vERR_CNT > 0)
                {
                    isLoad = false;
                    pPM.PromptText = string.Format("Excel Uploading Error : {0} :: {1}", vERR_CNT, vErrorMessage);
                }
                else
                {
                    isLoad = true;
                }
            }
            catch (System.Exception ex)
            {
                DisposeXL();

                mAppInterface.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return isLoad;
        }

        #endregion;
    }
}
