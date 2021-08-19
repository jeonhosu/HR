using System;
using ISCommonUtil;

namespace HRMF0716
{
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        // ��Ʈ�� ����.
        private string mTargetSheet = "PRINT";
        private string mSourceSheet1 = "SOURCE1";
        private string mSourceSheet2 = "SOURCE2";

        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // ù ������ üũ.

        // �μ�� ���ο� �հ�.
        private int mCopyLineSUM = 0;

        // �μ� 1���� �ִ� �μ�����.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 12;
        private int mCopy_EndRow = 36;

        private int m1stLastRow = 49;       //ù�� ���� �μ� ����.

        private int mPrintingLastRow = 37;  //���� �μ� ���� ���� ����.

        private int mPromptRow = 1;
        private int mCurrentRow = 2;       //���� �μ�Ǵ� row ��ġ.

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

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
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
        {// ���ϸ� �ڿ� �Ϸù�ȣ ����.
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


        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {// ���ڿ� ���� üũ �� �ش� �� ����.
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
        {// ���� ���� üũ �� �ش� �� ����.
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
        {// ��¥ ���� üũ �� �ش� �� ����.
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

        #region ----- Header Write Method ----

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID, Object pSTANDARD_DATE)
        {// ��� �μ�.
            int vCurrentCol = 0;
            int vTotalRow = pGRID.GridAdvExColElement[1].HeaderElement.Count;
            int vTotalCol = pGRID.ColCount;

            object vValue = null;
            string vVisible_YN = "0";

            try
            {
                if (vTotalCol > 0)
                {
                    #region ----- Write Page Copy(SourceSheet => TargetSheet) ----
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);
                }
                    #endregion;

                mCurrentRow = 1;
                vCurrentCol = 1;

                //�μ� �Ͻ� ǥ��
                mPrinting.XLSetCell(mCurrentRow, vCurrentCol, string.Format("Print Datetime : {0}", pSTANDARD_DATE));

                mCurrentRow = mCurrentRow + 1;


                mCopy_EndCol = vCurrentCol;  // copy ���� ����.
                vCurrentCol = 0;

                for (int r = 1; r < vTotalRow; r--)
                {
                    for (int c = 0; c < vTotalCol; c++)
                    {// ������Ʈ ǥ��.
                        vVisible_YN = iString.ISNull(pGRID.GridAdvExColElement[c].Visible, "0");

                        if (vVisible_YN == "1")
                        {
                            vValue = pGRID.GridAdvExColElement[c].HeaderElement[r].TL1_KR;


                            vCurrentCol = vCurrentCol + 1;
                            mPrinting.XLSetCell(mCurrentRow, vCurrentCol, vValue);
                            mPrinting.XL_LineDraw_Right(mCurrentRow, vCurrentCol, vCurrentCol, 1);   //������Ʈ ���� �׸���.

                        }
                    }
                    vCurrentCol = 0;
                    mCurrentRow = mCurrentRow + 1;
                }

                //���� �׸��� ����.
                mPrinting.XL_LineDraw_TopBottom(mPromptRow + 1, 1, vCurrentCol, 2);
                mPrinting.XL_LineDraw_Left(mPromptRow + 1, 1, 1, 2);
                mPrinting.XL_LineDraw_Right(mPromptRow + 1, vCurrentCol, vCurrentCol, 2);
                //���� �׸��� ����.

                mPrinting.XLCellAlignmentHorizontal(mPromptRow + 1, 1, mPromptRow, vCurrentCol, "C");
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Excel Line Wirte Methods ----

        public int LineWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID)
        {// ���� ȣ��Ǵ� �κ�.
            string vMessage = string.Empty;
            string vVisible_YN = "0";

            int vCurrentCol = 0;
            int vTotalRow = pGRID.RowCount;
            int vTotalCol = pGRID.ColCount;
            decimal vNumberValue = 0;

            object vDecimalDigit = 0;
            object vColumnType = null;
            object vValue = null;
            string vPrintValue = null;

            try
            {
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.                

                if (vTotalRow > 0)
                {
                    //#region ----- Write Page Copy(SourceSheet => TargetSheet) ----
                    //// ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    //mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                    //#endregion;

                    mPrinting.XLCellAlignmentHorizontal(mPromptRow, 1, mPromptRow, vCurrentCol, "C");
                    mCopy_EndCol = vCurrentCol;  // copy ���� ����.

                    for (int r = 0; r < vTotalRow; r++)
                    {//Row
                        for (int c = 0; c < vTotalCol; c++)
                        {//Col
                            vVisible_YN = iString.ISNull(pGRID.GridAdvExColElement[c].Visible, "0");
                            if (vVisible_YN == "1")
                            {
                                vCurrentCol = vCurrentCol + 1;
                                vValue = pGRID.GetCellValue(r, c);
                                vColumnType = pGRID.GridAdvExColElement[c].ColumnType;
                                vDecimalDigit = pGRID.GridAdvExColElement[c].DecimalDigits;
                                if (iString.ISNull(vColumnType) == "NumberEdit")
                                {
                                    try
                                    {
                                        vNumberValue = iString.ISDecimaltoZero(vValue);
                                        if (iString.ISNumtoZero(vDecimalDigit) > 0)
                                        {
                                            vPrintValue = string.Format("{0:###,###,###,###,###,###,###,###,###.####}", vNumberValue);
                                        }
                                        else
                                        {
                                            vPrintValue = string.Format("{0:###,###,###,###,###,###,###,###,##0}", vNumberValue);
                                        }
                                    }
                                    catch
                                    {
                                        vPrintValue = iString.ISNull(vValue);
                                    }
                                    mPrinting.XLCellAlignmentHorizontal(mCurrentRow, vCurrentCol, mCurrentRow, vCurrentCol, "R");
                                }
                                else if (iString.ISNull(vColumnType) == "DateTimeEdit")
                                {
                                    try
                                    {
                                        if (iString.ISNumtoZero(vDecimalDigit) > 0)
                                        {
                                            vPrintValue = string.Format("{0}", iDate.ISGetDate(vValue).ToShortDateString());
                                        }
                                        else
                                        {
                                            vPrintValue = string.Format("{0}", iDate.ISGetDate(vValue).ToShortDateString());
                                        }
                                    }
                                    catch
                                    {
                                        vPrintValue = iString.ISNull(vValue);
                                    }

                                    if (vPrintValue == "0001-01-01")
                                    {
                                        vPrintValue = string.Empty;
                                    }
                                    mPrinting.XLCellAlignmentHorizontal(mCurrentRow, vCurrentCol, mCurrentRow, vCurrentCol, "R");
                                }
                                else
                                {
                                    vPrintValue = iString.ISNull(vValue);
                                }
                                mPrinting.XLSetCell(mCurrentRow, vCurrentCol, vPrintValue);
                            }
                            vMessage = String.Format("Writing - [{0}/{1}]", r, vTotalRow);
                            mAppInterface.OnAppMessageEvent(vMessage);
                            System.Windows.Forms.Application.DoEvents();
                        }
                        vCurrentCol = 0;
                        mCurrentRow = mCurrentRow + 1;
                    }
                    mPrinting.XLColumnAutoFit(1, 1, mCurrentRow, mCopy_EndCol);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount)
        {
            int iDefaultEndRow = 1;
            if (mPageNumber == 1)
            {
                if (pPageRowCount == m1stLastRow)
                { // pPrintingLine : ���� ��µ� ��.
                    mIsNewPage = true;
                    iDefaultEndRow = mCopy_EndRow - m1stLastRow - 1;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow + iDefaultEndRow);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
            else
            {
                if (pPageRowCount == mPrintingLastRow)
                { // pPrintingLine : ���� ��µ� ��.
                    mIsNewPage = true;
                    iDefaultEndRow = mCopy_EndRow - mPrintingLastRow - 1;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCurrentRow + iDefaultEndRow);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //������ ActiveSheet�� ������ ����  ������ ����
        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            // page�� ǥ��.
            mPageNumber = mPageNumber + 1;

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(pActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.

            return vPasteEndRow;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            if (iString.ISNull(pSaveFileName) == string.Empty)
            {
                return;
            }

            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
