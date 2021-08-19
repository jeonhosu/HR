using System;
using ISCommonUtil;

namespace HRMF0515
{
    public class XLoading
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private string mMessageError = string.Empty;

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;
        
        private XL.XLPrint mImport = null;

        private string mXLOpenFileName = string.Empty;

        private int mCountROW = 0;    //Excel Active Sheet Row Count
        private int mCountCOLUMN = 0; //Excel Active Sheet Column Count

        private int mStartRowRead = 0; //읽을 시작 행

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

        public int CountROW
        {
            get
            {
                return mCountROW;
            }
            set
            {
                mCountROW = value;
            }
        }

        public int CountCOLUMN
        {
            get
            {
                return mCountCOLUMN;
            }
            set
            {
                mCountCOLUMN = value;
            }
        }

        public int ReadRow
        {
            get
            {
                return mStartRowRead;
            }
            set
            {
                mStartRowRead = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLoading()
        {
            mImport = new XL.XLPrint();
        }

        public XLoading(InfoSummit.Win.ControlAdv.ISAppInterfaceAdv pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mAppInterface = pAppInterface;
            mMessageAdapter = pMessageAdapter;

            mImport = new XL.XLPrint();
        }

        #endregion;

        #region ----- XLDispose -----

        public void DisposeXL()
        {
            mImport.XLOpenFileClose();
            mImport.XLClose();
        }

        #endregion;

        #region ----- XL File Open -----

        public bool OpenXL()
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mImport.XLFileOpen(mXLOpenFileName);

                mImport.XLActiveSheet(1);

                mCountROW = mImport.CountROW + 1;
                mCountCOLUMN = mImport.CountCOLUMN + 1;
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

        public bool LoadXL(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {
            string vMessage = string.Empty;

            bool isLoad = false;
            System.Type vType = null;
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDate = System.DateTime.Now;

            string visString1 = string.Empty; //급사여년월
            string visString2 = string.Empty; //급상여구분
            string visString4 = string.Empty; //사번

            int vCountXLRow = 0;
            int vCountXLColumn = 0;

            int vCountGridRow = 0;
            int vCountGridColumn = 0;
            int vStartRow = mStartRowRead;

            try
            {
                vCountXLRow = mCountROW;
                vCountXLColumn = mCountCOLUMN;

                for (int vRow = vStartRow; vRow < vCountXLRow; vRow++)
                {
                    vObject = mImport.XLGetCell(vRow, 1);
                    vConvertString = iString.ISNull(vObject);
                    visString1 = vConvertString.Trim();

                    vObject = mImport.XLGetCell(vRow, 2);
                    vConvertString = iString.ISNull(vObject);
                    visString2 = vConvertString.Trim();

                    vObject = mImport.XLGetCell(vRow, 4);
                    vConvertString = iString.ISNull(vObject);
                    visString4 = vConvertString.Trim();

                    if (string.IsNullOrEmpty(visString1) != true && string.IsNullOrEmpty(visString2) != true && string.IsNullOrEmpty(visString4) != true)
                    {
                        pAdapter.AddUnder();
                        for (int vCol = 1; vCol < vCountXLColumn; vCol++)
                        {
                            vType = mImport.XLGetType(vRow, vCol);

                            vObject = mImport.XLGetCell(vRow, vCol);

                            if (vType != null && vObject != null)
                            {
                                if (vType.Name == "String")
                                {
                                    vConvertString = ConvertString(vObject);
                                    vConvertString = vConvertString.Trim();
                                    pAdapter.CurrentRow[vCountGridColumn] = vConvertString;
                                }
                                else if (vType.Name == "Decimal")
                                {
                                    vConvertDecimal = ConvertDecimal(vObject);
                                    pAdapter.CurrentRow[vCountGridColumn] = vConvertDecimal;
                                }
                                else if (vType.Name == "Double")
                                {
                                    vConvertDecimal = ConvertDouble(vObject);
                                    pAdapter.CurrentRow[vCountGridColumn] = vConvertDecimal;
                                }
                                else if (vType.Name == "DateTime")
                                {
                                    vConvertDate = ConvertDate(vObject);
                                    pAdapter.CurrentRow[vCountGridColumn] = vConvertDate;
                                }
                            }

                            vCountGridColumn++;
                        }
                    }

                    vCountGridRow++;
                    vCountGridColumn = 0;

                    vMessage = string.Format("{0:D4}/{1:D4}", vRow, (vCountXLRow - 1));
                    mAppInterface.OnAppMessage(vMessage);
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
