using System;
using ISCommonUtil;

namespace HRMF0401
{
    public class XL_Upload
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();


        private string mMessageError = string.Empty;

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;
        
        private XL.XLPrint mExcel_Upload = null;

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

                mExcel_Upload.XLActiveSheet(1);

                mTotalROW = mExcel_Upload.CountROW + 1;
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
            System.DateTime vDate = System.DateTime.Now;

            try
            {
                if (pObject != null)
                {
                    isConvert = pObject is double;
                    if (isConvert == true)
                    {
                        double isConvertDouble = (double)pObject;
                        vDate = System.DateTime.FromOADate(isConvertDouble);
                        vTextDateTimeShort = vDate.ToString("yyyy-MM-dd", null);
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

        public bool LoadXL_P(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pStartRow)
        {
            string vMessage = string.Empty;

            bool isLoad = false;
            System.Type vType = null;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vConvertDate = new DateTime();

            object vPERSON_NUM = string.Empty;

            int vADRow = 0;
            int vADCol = 0;

            mTotalCOLUMN = pAdapter.OraSelectData.Columns.Count - 2;  //-2의미는 아답터에 기본 2COLUMN이 포함되어 있으므로(ROWKEYID, ROW STATUS) //

            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {
                    pAdapter.AddUnder();
                    //KEY값에 해당하는 셀에 DATA가 있을 경우만 INSERT를 처리해야 하므로//
                    vType = pAdapter.CurrentRow.Table.Columns["PERSON_NUM"].DataType;
                    if (vType.Name == "String")
                    {
                        vPERSON_NUM = mExcel_Upload.XLGetCell(vRow, 1);  //KEY 값.
                    }
                    else
                    {
                        vPERSON_NUM = string.Empty;
                        pAdapter.Delete();
                    }
                    if (iString.ISNull(vPERSON_NUM).Trim() != string.Empty)  //자산명 있을 경우만 처리.
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
                                    if (vConvertDate == iDate.ISGetDate("-"))
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
                    else
                    {
                        pAdapter.Delete();
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

        public bool LoadXL_M(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pStartRow)
        {
            string vMessage = string.Empty;

            bool isLoad = false;
            System.Type vType = null;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vConvertDate = new DateTime();

            object vPERSON_NUM = string.Empty;

            int vADRow = 0;
            int vADCol = 0;

            mTotalCOLUMN = pAdapter.OraSelectData.Columns.Count - 2;  //-2의미는 아답터에 기본 2COLUMN이 포함되어 있으므로(ROWKEYID, ROW STATUS) //

            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {
                    pAdapter.AddUnder();
                    //KEY값에 해당하는 셀에 DATA가 있을 경우만 INSERT를 처리해야 하므로//
                    vType = pAdapter.CurrentRow.Table.Columns["PERSON_NUM"].DataType;
                    if (vType.Name == "String")
                    {
                        vPERSON_NUM = mExcel_Upload.XLGetCell(vRow, 1);  //KEY 값.
                    }
                    else
                    {
                        vPERSON_NUM = string.Empty;
                        pAdapter.Delete();
                    }
                    if (iString.ISNull(vPERSON_NUM).Trim() != string.Empty)  //자산명 있을 경우만 처리.
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
                                    if (vConvertDate == iDate.ISGetDate("-"))
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
                    else
                    {
                        pAdapter.Delete();
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

    }
}
