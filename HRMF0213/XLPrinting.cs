using System;

namespace HRMF0213
{
    /// <summary>
    /// XLPrint Class�� �̿��� Report�� ���� 
    /// </summary>
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISGridAdvEx mGridAdvEx;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar1;
        private InfoSummit.Win.ControlAdv.ISProgressBar mProgressBar2;

        private XL.XLPrint mPrinting = null;

        // ��Ʈ�� ����.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "SOURCE1";

        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        // �μ� 1���� �ִ� �μ�����.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 15;
        private int mCopy_EndRow = 20;

        #endregion;

        #region ----- Property -----

        /// <summary>
        /// ��� Error Message ���
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        /// <summary>
        /// Message ����� Grid
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISGridAdvEx MessageGridEx
        {
            set
            {
                mGridAdvEx = value;
            }
        }

        /// <summary>
        /// ��ü Data ���� ProgressBar
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar1
        {
            set
            {
                mProgressBar1 = value;
            }
        }

        /// <summary>
        /// Page�� Data ���� ProgressBar
        /// </summary>
        public InfoSummit.Win.ControlAdv.ISProgressBar ProgressBar2
        {
            set
            {
                mProgressBar2 = value;
            }
        }

        /// <summary>
        /// Ope�� Excel File �̸�
        /// </summary>
        public string OpenFileNameExcel
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting()
        {
            mPrinting = new XL.XLPrint();
        }

        #endregion;

        #region ----- Interior Use Methods ----

        #region ----- MessageGrid Methods ----

        private void MessageGrid(string pMessage)
        {
            int vCountRow = mGridAdvEx.RowCount;
            vCountRow = vCountRow + 1;
            mGridAdvEx.RowCount = vCountRow;

            int vCurrentRow = vCountRow - 1;

            mGridAdvEx.SetCellValue(vCurrentRow, 0, pMessage);

            mGridAdvEx.CurrentCellMoveTo(vCurrentRow, 0);
            mGridAdvEx.Focus();
            mGridAdvEx.CurrentCellActivate(vCurrentRow, 0);
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

        #endregion;

        #region ----- Dispose -----

        public void Dispose()
        {
            mPrinting.XLOpenFileClose();
            mPrinting.XLClose();
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

        #region ----- Print Content Write Methods ----

        private void XLContentWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pIndexRow)
        {
            try
            {
                int vPrintRow = 14;
                int vPrintCol = 1;

                mPrinting.XLActiveSheet("SOURCE1");

                // �⺻ ����1
                int vIndexDataColumn1  = pGrid.GetColumnToIndex("NAME");            // ����
                int vIndexDataColumn3  = pGrid.GetColumnToIndex("DEPT_NAME");       // �μ�    
                int vIndexDataColumn6  = pGrid.GetColumnToIndex("PERSON_NUM");      // ���  
                int vIndexDataColumn83 = pGrid.GetColumnToIndex("FLOOR_NAME");      // �۾���

                // �ʱ�ȭ //
                //�μ�
                mPrinting.XLSetCell(vPrintRow, vPrintCol, string.Empty);

                //����
                vPrintRow = vPrintRow + 1;
                mPrinting.XLSetCell(vPrintRow, vPrintCol, string.Empty);

                //���
                vPrintRow = vPrintRow + 1;
                mPrinting.XLSetCell(vPrintRow, vPrintCol, string.Empty);

                //�۾���
                vPrintRow = vPrintRow + 1;
                mPrinting.XLSetCell(vPrintRow, vPrintCol, string.Empty);

                vPrintRow = 14;

                // ���� ������ �μ� //
                //�μ�
                mPrinting.XLSetCell(vPrintRow, vPrintCol, pGrid.GetCellValue(pIndexRow, vIndexDataColumn3));

                //����
                vPrintRow = vPrintRow + 1;
                mPrinting.XLSetCell(vPrintRow, vPrintCol, pGrid.GetCellValue(pIndexRow, vIndexDataColumn1));

                //���
                vPrintRow = vPrintRow + 1;
                mPrinting.XLSetCell(vPrintRow, vPrintCol, pGrid.GetCellValue(pIndexRow, vIndexDataColumn6));

                //�۾���
                vPrintRow = vPrintRow + 1;
                mPrinting.XLSetCell(vPrintRow, vPrintCol, pGrid.GetCellValue(pIndexRow, vIndexDataColumn83));

            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Excel Open and Close ----

        public void XLOpenClose()
        {
            mPrinting.XLOpenFileClose();

            XLFileOpen();
        }
        #endregion;

        #region ----- Excel Wirte Methods ----

        public void XLWirte(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pTerritory
                            , string pImageName, int vWriteRow, int vWriteCol
                            , System.Drawing.SizeF pSize, System.Drawing.PointF pPoint)
        {
            string vMessageText = string.Empty;
            int pPicture = 2;  //�μ� ��Ŀ� �ִ� �̹������� +1(�������)�� ��;

            try
            {     
                // ������� �μ� //
                XLContentWrite(pGrid, pRow);

                //������� �μ�
                //System.Drawing.SizeF vSize = new System.Drawing.SizeF(95.2283F, 110.99701F);
                //System.Drawing.PointF vPoint = new System.Drawing.PointF(34F, 53F);

                // ������� �μ� //
                //mPrinting.XLBarCode(3, , pSize, pPoint);

                mPrinting.XLActiveSheet("SOURCE1");
                try
                {
                    mPrinting.XLDeleteBarCode(pPicture);
                    mPrinting.XLBarCode(pPicture, pImageName, pSize, pPoint);
                }
                catch
                {

                }

                //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
                mPageNumber = CopyAndPaste(mPrinting, mSourceSheet1, vWriteRow, vWriteCol);


                mPrinting.XLActiveSheet("SOURCE1");
                mPrinting.XLDeleteBarCode(pPicture);                
                //mPrinting.XLBarCode(pPicture, string.Empty, pSize, pPoint);
                //-------------------------------------------------------------------------------------------------------
                // ������ ���� ���� �κ�
                // (SourceTab1�� ������ ��� -> Destination�� ���� -> SourceTab1 ������ ���� ��, ���� ������ ��� 
                //-------------------------------------------------------------------------------------------------------

            }
            catch
            {
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow, int pPasteStartCol)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;
            int vPasteEndCol = pPasteStartCol + mCopy_EndCol;

            // page�� ǥ��.
            mPageNumber = mPageNumber + 1;
            //XLPageNumber(pActiveSheet, mPageNumber);

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(pActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, pPasteStartCol, vPasteEndRow, vPasteEndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.

            return vPasteEndRow;

        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            try
            {
                mPrinting.XLPrinting(pPageSTART, pPageEND);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            if (pSaveFileName == string.Empty)
            {
                return;
            }

            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
