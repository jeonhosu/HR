using System;
using ISCommonUtil;

namespace HRMF0106
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
                vString = iString.ISNull(pObject);
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
                    else
                    {
                        vConvertDecimal = 0;
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
                    //KEY���� �ش��ϴ� ���� DATA�� ���� ��츸 INSERT�� ó���ؾ� �ϹǷ�//
                    vType = pAdapter.CurrentRow.Table.Columns["PERSON_NUM"].DataType;
                    if (vType.Name == "String")
                    {
                        vObject = mExcel_Upload.XLGetCell(vRow, 2);  //�����ȣ.
                        vPERSON_NUM = iString.ISNull(vObject);
                    }
                    else 
                    {
                        vPERSON_NUM = string.Empty;
                        pAdapter.Delete();
                    }
                    if (iString.ISNull(vPERSON_NUM) != string.Empty)  //�����ȣ�� ���� ��츸 ó��.
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

        public bool LoadXL(InfoSummit.Win.ControlAdv.ISDataCommand pCMD, int pStartRow, InfoSummit.Win.ControlAdv.ISProgressBar pPB, InfoSummit.Win.ControlAdv.ISPrompt pPM, object pSTD_DATE)
        {
            string vMessage = string.Empty;
             
            mExcel_Upload.XLActiveSheet(1);
            mTotalROW = mExcel_Upload.CountROW + 1; 
            
            bool isLoad = false; 

            object vObject = null; 
            decimal vSTART_AMOUNT = 0;
            decimal vEND_AMOUNT = 0;
            decimal vCOL3 = 0;
            decimal vCOL4 = 0;
            decimal vCOL5 = 0;
            decimal vCOL6 = 0;
            decimal vCOL7 = 0;
            decimal vCOL8 = 0;
            decimal vCOL9 = 0;
            decimal vCOL10 = 0;
            decimal vCOL11 = 0;
            decimal vCOL12 = 0;
            decimal vCOL13 = 0;

            int vADRow = 0;
            int vERR_CNT = 0;

            try
            {
                for (int vRow = pStartRow; vRow < mTotalROW; vRow++)
                {                    
                    //KEY���� �ش��ϴ� ���� DATA�� ���� ��츸 INSERT�� ó���ؾ� �ϹǷ�//
                    vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 1)).Replace("-", "");  //���޿��� START  
                    if(iString.ISDecimal(vObject) == true)
                    {
                        vSTART_AMOUNT = iString.ISDecimaltoZero(vObject); 
                    }
                    else
                    {
                        vSTART_AMOUNT = -1;
                    }
                    vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 2)).Replace("-", "");  //���޿��� END  
                    if (iString.ISDecimal(vObject) == true)
                    {
                        vEND_AMOUNT = iString.ISDecimaltoZero(vObject);
                    }
                    else
                    {
                        vEND_AMOUNT = -1;
                    } 

                    //�ξ簡��.
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 3)).Replace("-", "");  //�ξ簡��1.
                        vCOL3 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL3 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 4)).Replace("-", "");  //�ξ簡��2
                        vCOL4 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL4 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 5)).Replace("-", ""); 
                        vCOL5 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL5 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 6)).Replace("-", ""); 
                        vCOL6 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL6 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 7)).Replace("-", "");  
                        vCOL7 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL7 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 8)).Replace("-", "");  
                        vCOL8 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL8 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 9)).Replace("-", ""); 
                        vCOL9 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL9 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 10)).Replace("-", "");  
                        vCOL10 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL10 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 11)).Replace("-", "");  
                        vCOL11 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL11 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 12)).Replace("-", "");   
                        vCOL12 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL12 = 0;
                    }
                    try
                    {
                        vObject = iString.ISNull(mExcel_Upload.XLGetCell(vRow, 13)).Replace("-", "");  
                        vCOL13 = iString.ISDecimaltoZero(vObject);
                    }
                    catch
                    {
                        vCOL13 = 0;
                    }

                    if (vSTART_AMOUNT > 0) 
                    {
                        try
                        {                           
                            pCMD.SetCommandParamValue("P_START_AMOUNT", vSTART_AMOUNT);
                            pCMD.SetCommandParamValue("P_END_AMOUNT", vEND_AMOUNT);
                            pCMD.SetCommandParamValue("P_DED_PERSON_1",vCOL3);
                            pCMD.SetCommandParamValue("P_DED_PERSON_2",vCOL4);
                            pCMD.SetCommandParamValue("P_DED_PERSON_3",vCOL5);
                            pCMD.SetCommandParamValue("P_DED_PERSON_4",vCOL6);
                            pCMD.SetCommandParamValue("P_DED_PERSON_5",vCOL7);
                            pCMD.SetCommandParamValue("P_DED_PERSON_6",vCOL8);
                            pCMD.SetCommandParamValue("P_DED_PERSON_7",vCOL9);
                            pCMD.SetCommandParamValue("P_DED_PERSON_8",vCOL10);
                            pCMD.SetCommandParamValue("P_DED_PERSON_9",vCOL11);
                            pCMD.SetCommandParamValue("P_DED_PERSON_10",vCOL12);
                            pCMD.SetCommandParamValue("P_DED_PERSON_11",vCOL13);
                            pCMD.SetCommandParamValue("P_STD_DATE", pSTD_DATE);
                            pCMD.ExecuteNonQuery(); 
                            if (iString.ISNull(pCMD.GetCommandParamValue("O_STATUS")) == "F")
                            {
                                vMessage = iString.ISNull(pCMD.GetCommandParamValue("O_MESSAGE"));
                                vERR_CNT++;
                                pPM.PromptText = string.Format("Imporing :: {0}-{1} *** {2}({3} ** Error : {4})", vADRow, mTotalROW, vSTART_AMOUNT, vEND_AMOUNT, vMessage);
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

                    pPB.BarFillPercent = (Convert.ToSingle(vADRow + pStartRow) / Convert.ToSingle(mTotalROW)) * 100F;
                    pPM.PromptText = string.Format("Imporing Counting :: {0} / {1} *** Amount :: {2} ~ {3}", vADRow, mTotalROW, vSTART_AMOUNT, vEND_AMOUNT);

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
