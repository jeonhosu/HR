using System;
using ISCommonUtil;

namespace HRMF0303
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
             
            mExcel_Upload.XLActiveSheet(1);
            mTotalROW = mExcel_Upload.CountROW + 1; 

            bool isLoad = false;
             
            object vObject = null;
            object vNAME = string.Empty;
            object vPERSON_NUM = string.Empty;
            object vWORK_DATE = string.Empty; 

            int vADRow = 0;
            int vERR_CNT = 0;

            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {                    
                    //KEY값에 해당하는 셀에 DATA가 있을 경우만 INSERT를 처리해야 하므로//
                    //vObject = mExcel_Upload.XLGetCell(vRow, 1);  //근무일자.
                    //vWORK_DATE = ConvertDate(vObject);
                    vPERSON_NUM = mExcel_Upload.XLGetCell(vRow, 2);
                    vNAME = string.Empty;

                    if (iString.ISNull(vPERSON_NUM) != string.Empty)  //사원번호가 있을 경우만 처리.
                    {
                        try
                        {
                            vNAME = mExcel_Upload.XLGetCell(vRow, 3);
                              
                            pCMD.SetCommandParamValue("P_PERSON_NUM", vPERSON_NUM);
                            pCMD.SetCommandParamValue("P_NAME", vNAME);
                            pCMD.SetCommandParamValue("P_WORK_TYPE_NAME", mExcel_Upload.XLGetCell(vRow, 1));
                            //pCMD.SetCommandParamValue("P_WORK_TYPE", mExcel_Upload.XLGetCell(vRow, 36));
                            pCMD.SetCommandParamValue("P_D01", mExcel_Upload.XLGetCell(vRow, 5));
                            pCMD.SetCommandParamValue("P_D02", mExcel_Upload.XLGetCell(vRow, 6));
                            pCMD.SetCommandParamValue("P_D03", mExcel_Upload.XLGetCell(vRow, 7));
                            pCMD.SetCommandParamValue("P_D04", mExcel_Upload.XLGetCell(vRow, 8));
                            pCMD.SetCommandParamValue("P_D05", mExcel_Upload.XLGetCell(vRow, 9));
                            pCMD.SetCommandParamValue("P_D06", mExcel_Upload.XLGetCell(vRow, 10));
                            pCMD.SetCommandParamValue("P_D07", mExcel_Upload.XLGetCell(vRow, 11));
                            pCMD.SetCommandParamValue("P_D08", mExcel_Upload.XLGetCell(vRow, 12));
                            pCMD.SetCommandParamValue("P_D09", mExcel_Upload.XLGetCell(vRow, 13));
                            pCMD.SetCommandParamValue("P_D10", mExcel_Upload.XLGetCell(vRow, 14));
                            pCMD.SetCommandParamValue("P_D11", mExcel_Upload.XLGetCell(vRow, 15));
                            pCMD.SetCommandParamValue("P_D12", mExcel_Upload.XLGetCell(vRow, 16));
                            pCMD.SetCommandParamValue("P_D13", mExcel_Upload.XLGetCell(vRow, 17));
                            pCMD.SetCommandParamValue("P_D14", mExcel_Upload.XLGetCell(vRow, 18));
                            pCMD.SetCommandParamValue("P_D15", mExcel_Upload.XLGetCell(vRow, 19));
                            pCMD.SetCommandParamValue("P_D16", mExcel_Upload.XLGetCell(vRow, 20));
                            pCMD.SetCommandParamValue("P_D17", mExcel_Upload.XLGetCell(vRow, 21));
                            pCMD.SetCommandParamValue("P_D18", mExcel_Upload.XLGetCell(vRow, 22));
                            pCMD.SetCommandParamValue("P_D19", mExcel_Upload.XLGetCell(vRow, 23));
                            pCMD.SetCommandParamValue("P_D20", mExcel_Upload.XLGetCell(vRow, 24));
                            pCMD.SetCommandParamValue("P_D21", mExcel_Upload.XLGetCell(vRow, 25));
                            pCMD.SetCommandParamValue("P_D22", mExcel_Upload.XLGetCell(vRow, 26));
                            pCMD.SetCommandParamValue("P_D23", mExcel_Upload.XLGetCell(vRow, 27));
                            pCMD.SetCommandParamValue("P_D24", mExcel_Upload.XLGetCell(vRow, 28));
                            pCMD.SetCommandParamValue("P_D25", mExcel_Upload.XLGetCell(vRow, 29));
                            pCMD.SetCommandParamValue("P_D26", mExcel_Upload.XLGetCell(vRow, 30));
                            pCMD.SetCommandParamValue("P_D27", mExcel_Upload.XLGetCell(vRow, 31));
                            pCMD.SetCommandParamValue("P_D28", mExcel_Upload.XLGetCell(vRow, 32));
                            pCMD.SetCommandParamValue("P_D29", mExcel_Upload.XLGetCell(vRow, 33));
                            pCMD.SetCommandParamValue("P_D30", mExcel_Upload.XLGetCell(vRow, 34));
                            pCMD.SetCommandParamValue("P_D31", mExcel_Upload.XLGetCell(vRow, 35)); 
                            pCMD.ExecuteNonQuery();
                            if (iString.ISNull(pCMD.GetCommandParamValue("O_STATUS")) == "F")
                            {
                                vMessage = iString.ISNull(pCMD.GetCommandParamValue("O_MESSAGE"));
                                pPM.PromptText = string.Format("Importing :: {0} / {1} *** Name : {2}({3} ** Error : {4})", vRow, (mTotalROW - pStartRow), vNAME, vPERSON_NUM, vMessage);
                                vERR_CNT++;
                                return false;
                            }
                        }
                        catch (Exception Ex)
                        {
                            DisposeXL();

                            vMessage = iString.ISNull(pCMD.GetCommandParamValue("O_MESSAGE"));
                            pPM.PromptText = string.Format("Importing :: {0} / {1} *** Name : {2}({3} ** Error : {4})", vRow, (mTotalROW - pStartRow), vNAME, vPERSON_NUM, vMessage);
                            vERR_CNT++;

                            mAppInterface.OnAppMessage(Ex.Message);
                            System.Windows.Forms.Application.DoEvents();
                            return false;
                        }
                    }
                    vADRow++;

                    pPB.BarFillPercent = (Convert.ToSingle(vADRow) / Convert.ToSingle(mTotalROW)) * 100F;
                    pPM.PromptText = string.Format("Importing :: {0} / {1} *** Name : {2}({3})", vRow, (mTotalROW - 1), vNAME, vPERSON_NUM);

                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();
                }
                if (vERR_CNT > 0)
                {
                    isLoad = false;
                    pPM.PromptText = string.Format("Importing Error : {0}", vMessage); 
                    mAppInterface.OnAppMessage(string.Format("Excel Uploading Error : {0}", vMessage));
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


        public bool LoadXL_Detail(InfoSummit.Win.ControlAdv.ISDataCommand pCMD, int pStartRow, InfoSummit.Win.ControlAdv.ISProgressBar pPB, InfoSummit.Win.ControlAdv.ISPrompt pPM)
        {
            string vMessage = string.Empty;

            mExcel_Upload.XLActiveSheet(1);
            mTotalROW = mExcel_Upload.CountROW + 1;

            bool isLoad = false;

            object vObject = null;
            object vWORK_TYPE_NAME = string.Empty;
            object vWORK_TYPE = string.Empty; 

            int vADRow = 0;
            int vERR_CNT = 0;

            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {
                    //KEY값에 해당하는 셀에 DATA가 있을 경우만 INSERT를 처리해야 하므로//
                    //vObject = mExcel_Upload.XLGetCell(vRow, 1);  //근무일자.
                    //vWORK_DATE = ConvertDate(vObject);
                    vWORK_TYPE = mExcel_Upload.XLGetCell(vRow, 1); 

                    if (iString.ISNull(vWORK_TYPE) != string.Empty)  //사원번호가 있을 경우만 처리.
                    {
                        try
                        {
                            vWORK_TYPE_NAME = mExcel_Upload.XLGetCell(vRow, 2);
                             
                            pCMD.SetCommandParamValue("P_WORK_TYPE", vWORK_TYPE);
                            pCMD.SetCommandParamValue("P_WORK_TYPE_NAME", vWORK_TYPE_NAME); 
                            pCMD.SetCommandParamValue("P_D01", mExcel_Upload.XLGetCell(vRow, 3));
                            pCMD.SetCommandParamValue("P_D02", mExcel_Upload.XLGetCell(vRow, 4));
                            pCMD.SetCommandParamValue("P_D03", mExcel_Upload.XLGetCell(vRow, 5));
                            pCMD.SetCommandParamValue("P_D04", mExcel_Upload.XLGetCell(vRow, 6));
                            pCMD.SetCommandParamValue("P_D05", mExcel_Upload.XLGetCell(vRow, 7));
                            pCMD.SetCommandParamValue("P_D06", mExcel_Upload.XLGetCell(vRow, 8));
                            pCMD.SetCommandParamValue("P_D07", mExcel_Upload.XLGetCell(vRow, 9));
                            pCMD.SetCommandParamValue("P_D08", mExcel_Upload.XLGetCell(vRow, 10));
                            pCMD.SetCommandParamValue("P_D09", mExcel_Upload.XLGetCell(vRow, 11));
                            pCMD.SetCommandParamValue("P_D10", mExcel_Upload.XLGetCell(vRow, 12));
                            pCMD.SetCommandParamValue("P_D11", mExcel_Upload.XLGetCell(vRow, 13));
                            pCMD.SetCommandParamValue("P_D12", mExcel_Upload.XLGetCell(vRow, 14));
                            pCMD.SetCommandParamValue("P_D13", mExcel_Upload.XLGetCell(vRow, 15));
                            pCMD.SetCommandParamValue("P_D14", mExcel_Upload.XLGetCell(vRow, 16));
                            pCMD.SetCommandParamValue("P_D15", mExcel_Upload.XLGetCell(vRow, 17));
                            pCMD.SetCommandParamValue("P_D16", mExcel_Upload.XLGetCell(vRow, 18));
                            pCMD.SetCommandParamValue("P_D17", mExcel_Upload.XLGetCell(vRow, 19));
                            pCMD.SetCommandParamValue("P_D18", mExcel_Upload.XLGetCell(vRow, 20));
                            pCMD.SetCommandParamValue("P_D19", mExcel_Upload.XLGetCell(vRow, 21));
                            pCMD.SetCommandParamValue("P_D20", mExcel_Upload.XLGetCell(vRow, 22));
                            pCMD.SetCommandParamValue("P_D21", mExcel_Upload.XLGetCell(vRow, 23));
                            pCMD.SetCommandParamValue("P_D22", mExcel_Upload.XLGetCell(vRow, 24));
                            pCMD.SetCommandParamValue("P_D23", mExcel_Upload.XLGetCell(vRow, 25));
                            pCMD.SetCommandParamValue("P_D24", mExcel_Upload.XLGetCell(vRow, 26));
                            pCMD.SetCommandParamValue("P_D25", mExcel_Upload.XLGetCell(vRow, 27));
                            pCMD.SetCommandParamValue("P_D26", mExcel_Upload.XLGetCell(vRow, 28));
                            pCMD.SetCommandParamValue("P_D27", mExcel_Upload.XLGetCell(vRow, 29));
                            pCMD.SetCommandParamValue("P_D28", mExcel_Upload.XLGetCell(vRow, 30));
                            pCMD.SetCommandParamValue("P_D29", mExcel_Upload.XLGetCell(vRow, 31));
                            pCMD.SetCommandParamValue("P_D30", mExcel_Upload.XLGetCell(vRow, 32));
                            pCMD.SetCommandParamValue("P_D31", mExcel_Upload.XLGetCell(vRow, 33));
                            pCMD.ExecuteNonQuery();
                            if (iString.ISNull(pCMD.GetCommandParamValue("O_STATUS")) == "F")
                            {
                                vMessage = iString.ISNull(pCMD.GetCommandParamValue("O_MESSAGE"));
                                pPM.PromptText = string.Format("Importing :: {0} / {1} *** Work type Name : {2}({3} ** Error : {4})", vRow, (mTotalROW - pStartRow), vWORK_TYPE, vWORK_TYPE_NAME, vMessage);
                                vERR_CNT++;
                                return false;
                            }
                        }
                        catch (Exception Ex)
                        {
                            DisposeXL();

                            vMessage = iString.ISNull(pCMD.GetCommandParamValue("O_MESSAGE"));
                            pPM.PromptText = string.Format("Importing :: {0} / {1} *** Work type Name : {2}({3} ** Error : {4})", vRow, (mTotalROW - pStartRow), vWORK_TYPE, vWORK_TYPE_NAME, vMessage);
                            vERR_CNT++;

                            mAppInterface.OnAppMessage(Ex.Message);
                            System.Windows.Forms.Application.DoEvents();
                            return false;
                        }
                    }
                    vADRow++;

                    pPB.BarFillPercent = (Convert.ToSingle(vADRow) / Convert.ToSingle(mTotalROW)) * 100F;
                    pPM.PromptText = string.Format("Importing :: {0} / {1} *** Work type Name : {2}({3})", vRow, (mTotalROW - 1), vWORK_TYPE, vWORK_TYPE_NAME);

                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();
                }
                if (vERR_CNT > 0)
                {
                    isLoad = false;
                    pPM.PromptText = string.Format("Importing Error : {0}", vMessage);
                    mAppInterface.OnAppMessage(string.Format("Excel Uploading Error : {0}", vMessage));
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
