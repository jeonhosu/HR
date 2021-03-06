using System;
using ISCommonUtil;

namespace HRMF0393
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
                    //KEY???? ???????? ???? DATA?? ???? ?????? INSERT?? ???????? ??????//
                    vType = pAdapter.CurrentRow.Table.Columns["PERSON_NUM"].DataType;
                    if (vType.Name == "String")
                    {
                        vObject = mExcel_Upload.XLGetCell(vRow, 2);  //????????.
                        vPERSON_NUM = iString.ISNull(vObject);
                    }
                    else 
                    {
                        vPERSON_NUM = string.Empty;
                        pAdapter.Delete();
                    }
                    if (iString.ISNull(vPERSON_NUM) != string.Empty)  //?????????? ???? ?????? ????.
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

            DateTime vConvertDate = new DateTime();
            object vObject = null;
            object vNAME = string.Empty;
            object vPERSON_NUM = string.Empty; 

            int vADRow = 0;
            int vERR_CNT = 0;

            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {                    
                    //KEY???? ???????? ???? DATA?? ???? ?????? INSERT?? ???????? ??????//
                    vObject = mExcel_Upload.XLGetCell(vRow, 1);  //????????.
                    vPERSON_NUM = iString.ISNull(vObject);
                    vNAME = string.Empty;

                    if (iString.ISNull(vPERSON_NUM) != string.Empty)  //?????????? ???? ?????? ????.
                    {
                        try
                        {
                            vNAME = mExcel_Upload.XLGetCell(vRow, 2);

                            vObject = string.Empty;
                            vObject = mExcel_Upload.XLGetCell(vRow, 5);
                            vConvertDate = ConvertDate(vObject);

                            pCMD.SetCommandParamValue("P_PERSON_NUM", vPERSON_NUM);
                            pCMD.SetCommandParamValue("P_NAME", vNAME);
                            pCMD.SetCommandParamValue("P_DEPT_NAME", mExcel_Upload.XLGetCell(vRow, 3));
                            pCMD.SetCommandParamValue("P_POSITION", mExcel_Upload.XLGetCell(vRow, 4));
                            pCMD.SetCommandParamValue("P_CHECK_DATE", vConvertDate);
                            pCMD.SetCommandParamValue("P_WEEK_NAME", mExcel_Upload.XLGetCell(vRow, 6));
                            pCMD.SetCommandParamValue("P_IN_TIME", mExcel_Upload.XLGetCell(vRow, 7));
                            pCMD.SetCommandParamValue("P_OUT_TIME", mExcel_Upload.XLGetCell(vRow, 8));
                            pCMD.SetCommandParamValue("P_ATTEND", mExcel_Upload.XLGetCell(vRow, 9));
                            pCMD.SetCommandParamValue("P_HOUR", mExcel_Upload.XLGetCell(vRow, 10));
                            pCMD.SetCommandParamValue("P_ATTEND1", mExcel_Upload.XLGetCell(vRow, 11));
                            pCMD.SetCommandParamValue("P_HOUR1", mExcel_Upload.XLGetCell(vRow, 12));
                            pCMD.SetCommandParamValue("P_LATE", mExcel_Upload.XLGetCell(vRow, 13));
                            pCMD.SetCommandParamValue("P_OUT_SOON", mExcel_Upload.XLGetCell(vRow, 14));
                            pCMD.SetCommandParamValue("P_OT1", mExcel_Upload.XLGetCell(vRow, 15));
                            pCMD.SetCommandParamValue("P_OT2", mExcel_Upload.XLGetCell(vRow, 16));
                            pCMD.SetCommandParamValue("P_OT3", mExcel_Upload.XLGetCell(vRow, 17));
                            pCMD.SetCommandParamValue("P_SHIFT_CODE", mExcel_Upload.XLGetCell(vRow, 18));
                            pCMD.SetCommandParamValue("P_SHIFT_CODE1", mExcel_Upload.XLGetCell(vRow, 19));
                            pCMD.SetCommandParamValue("P_SHIFT_CODE2", mExcel_Upload.XLGetCell(vRow, 20)); 
                            pCMD.ExecuteNonQuery();
                            if (iString.ISNull(pCMD.GetCommandParamValue("O_STATUS")) == "F")
                            {
                                vMessage = iString.ISNull(pCMD.GetCommandParamValue("O_MESSAGE"));
                                vERR_CNT++;
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
                    pPM.PromptText = string.Format("{0}-{1}", vPERSON_NUM, vNAME);

                    vMessage = string.Format("Excel Uploading : {0:D4}/{1:D4}", vRow, (mTotalROW - 1));
                    mAppInterface.OnAppMessage(vMessage);
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                    System.Windows.Forms.Application.DoEvents();
                }
                if (vERR_CNT > 0)
                {
                    isLoad = false;
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
