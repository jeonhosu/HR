using System;
using System.Drawing;
using System.Text;
using System.Globalization;
using ISCommonUtil;

namespace PPMF0638
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;
        
        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        private int mPrintingLineSTART1 = 10;  //Header
        private int mPrintingLineSTART2 = 16; //Line

        private int mCopyLineSUM = 1;        //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ, ���� �� ����
        private int mIncrementCopyMAX = 49;  //����Ǿ��� ���� ����

        private int mCopyColumnSTART = 1; //����Ǿ�  �� �� ���� ��
        private int mCopyColumnEND = 71;  //������ ���õ� ��Ʈ�� ����Ǿ��� �� �� ��ġ

        private int mCountROW = 0;

        private int mChoiceXLSheet = 0;

        private int mImageIndex = 1;
        
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

        #region ----- Array Set 1 ----

        private void SetArray1(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[12];
            pXLColumn = new int[12];

            pGDColumn[0] = pTable.Columns.IndexOf("PO_TYPE_NAME");
            pGDColumn[1] = pTable.Columns.IndexOf("DISPLAY_NAME"); //���Ŵ����
            pGDColumn[2] = pTable.Columns.IndexOf("PO_DATE");
            pGDColumn[3] = pTable.Columns.IndexOf("PO_NO");
            pGDColumn[4] = pTable.Columns.IndexOf("SUPPLIER_SHORT_NAME");
            pGDColumn[5] = pTable.Columns.IndexOf("TOTAL_AMOUNT");
            pGDColumn[6] = pTable.Columns.IndexOf("ADDRESS_1");
            pGDColumn[7] = pTable.Columns.IndexOf("ADDRESS_2");
            pGDColumn[8] = pTable.Columns.IndexOf("TELEPHONE_NO");
            pGDColumn[9] = pTable.Columns.IndexOf("STEP_DESCRIPTION");
            pGDColumn[10] = pTable.Columns.IndexOf("TELEPHONE_NO");
            pGDColumn[11] = pTable.Columns.IndexOf("FAX_NO");

            pXLColumn[0] = 9;    //PO_TYPE_NAME
            pXLColumn[1] = 10;   //DISPLAY_NAME
            pXLColumn[2] = 59;   //PO_DATE
            pXLColumn[3] = 59;   //PO_NO
            pXLColumn[4] = 33;    //SUPPLIER_SHORT_NAME ����ó
            pXLColumn[5] = 55;   //��ü TOTAL�ݾ�
            pXLColumn[6] = 33;   //PAYMENT_METHOD_NAME
            pXLColumn[7] = 33;   //PAYMENT_TERM_NAME
            pXLColumn[8] = 33;    //REMARK , TELEPHONE_NO
            pXLColumn[9] = 54;   //�ݾ�
            pXLColumn[10] = 55;  //��ȭ����
            pXLColumn[11] = 33;  //FAX_NO
        }

        #endregion;

        #region ----- Array Set 2 ----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[12];
            pXLColumn = new int[12];

            pGDColumn[0] = pGrid.GetColumnToIndex("ITEM_CODE");            //�����ڵ�
            pGDColumn[1] = pGrid.GetColumnToIndex("ITEM_DESCRIPTION");     //�����[ǰ��]
            pGDColumn[2] = pGrid.GetColumnToIndex("ITEM_UOM_CODE");        //UOM[����]
            pGDColumn[3] = pGrid.GetColumnToIndex("PO_QTY");               //��û���ַ�[����]
            pGDColumn[4] = pGrid.GetColumnToIndex("CURRENCY_CODE");        //��ȭ
            pGDColumn[5] = pGrid.GetColumnToIndex("ITEM_PRICE");           //�ܰ�
            pGDColumn[6] = pGrid.GetColumnToIndex("CURRENCY_CODE");        //��ȭ
            pGDColumn[7] = pGrid.GetColumnToIndex("ITEM_AMOUNT");          //�ݾ�
            pGDColumn[8] = pGrid.GetColumnToIndex("DELIVERY_REQ_DATE");    //����䱸��[������]
            pGDColumn[9] = pGrid.GetColumnToIndex("ITEM_SPECIFICATION");   //����԰�, add 14-10-14, by Ahn Sang Hyeon
            pGDColumn[10] = pGrid.GetColumnToIndex("COST_CENTER_DESC");   //�����μ�, add 14-10-14, by Ahn Sang Hyeon
            pGDColumn[11] = pGrid.GetColumnToIndex("REMARK");   //���

            pXLColumn[0] = 4;    //�����ڵ�
            pXLColumn[1] = 10;   //�����[ǰ��]
            pXLColumn[2] = 24;   //UOM[����]
            pXLColumn[3] = 27;   //��û���ַ�[����]
            pXLColumn[4] = 34;   //��ȭ
            pXLColumn[5] = 38;   //�ܰ�
            pXLColumn[6] = 45;   //��ȭ
            pXLColumn[7] = 48;   //�ݾ�
            pXLColumn[8] = 57;   //����䱸��[������]
            pXLColumn[9] = 18;   //����԰�, add 14-10-14, by Ahn Sang Hyeon
            pXLColumn[10] = 64;   //�����μ�, ADDED BY SHAN, 2017-07-05
            pXLColumn[11] = 63;   //�����μ�, ADDED BY SHAN, 2017-07-05

        }

        #endregion;

        #region ----- Convert decimal  Method ----

        private decimal ConvertNumber(string pStringNumber)
        {
            decimal vConvertDecimal = 0m;

            try
            {
                bool isNull = string.IsNullOrEmpty(pStringNumber);
                if (isNull != true)
                {
                    vConvertDecimal = decimal.Parse(pStringNumber);
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vConvertDecimal;
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

        #region ----- Write Method ----

        private void XLWrite(System.Data.DataRow pRow)
        {            
            
            
        }

        #endregion;

        #region ----- Line Write Method -----

        private int XLLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            System.DateTime vDateTime = new System.DateTime();
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //NO
                mCountROW++;
                vXLColumnIndex = 2;
                vConvertString = string.Format("{0:#0}", mCountROW);
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //�����ڵ�
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
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

                //�����[ǰ��]
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
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

                //����԰�, add 14-10-14, by Ahn Sang Hyeon
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
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

                //UOM[����]
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
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

                //��û���ַ�[����]
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //��ȭ
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
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

                //�ܰ�
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0.00}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //��ȭ
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
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

                //�ݾ�
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0.0000}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����䱸��[������]
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertDate(vObject, out vDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}-{1:D2}-{2:D2}", vDateTime.Year, vDateTime.Month, vDateTime.Day);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //��û�μ�
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
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

                //���
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
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

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite(System.Data.DataRow pRow)
        {
            //mPageNumber = 0;
            string vMessage = string.Empty;

            object vObject = null;
            object vVendor_Code = null;
            string vConvertString = string.Empty;
            System.DateTime vDateTime = new System.DateTime();
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                vVendor_Code = pRow["VENDOR_CODE"];
                IsConvert = IsConvertString(vVendor_Code, out vConvertString);

                int vIndexBarCodeImage = mPrinting.CountBarCodeImage;
                int vCountBarCodeImage = mPrinting.CountBarCodeImage;
                 
                mPrinting.CountBarCodeImage = 1;
                for (int vRow = 1; vRow < vCountBarCodeImage; vRow++)
                {
                    mPrinting.XLDeleteBarCode(vIndexBarCodeImage);
                    vIndexBarCodeImage--;
                }
                mPrinting.CountBarCodeImage = 1;

                if (vConvertString == "01004")
                {
                    mPrinting.XLActiveSheet(vConvertString);

                    /////////////////////���ڵ� ����/////////////////////// /////
                    Syncfusion.Pdf.Barcode.PdfCode128BBarcode barcode = new Syncfusion.Pdf.Barcode.PdfCode128BBarcode();

                    //barcode = Syncfusion.Pdf.Barcode.QRCodeVersion.Auto;
                    barcode.EncodeStartStopSymbols = true;
                    barcode.EnableCheckDigit = true;
                    barcode.ShowCheckDigit = true;
                    barcode.TextDisplayLocation = Syncfusion.Pdf.Barcode.TextLocation.None;
                    barcode.BarHeight = 90;
                    barcode.NarrowBarWidth = 1;
                    barcode.BarcodeToTextGapHeight = 20;

                    //Set the barcode text
                    barcode.Text = string.Format("{0}", pRow["PACKING_BOX_NO"]);

                    //Export the barcode as image
                    Image img = barcode.ToImage();

                    string vPath = System.Environment.CurrentDirectory + @"\BARCODE.jpg";
                    img.Save(vPath, System.Drawing.Imaging.ImageFormat.Png);

                     
                    System.Drawing.SizeF vSize = new System.Drawing.SizeF(430F, 60F);
                    System.Drawing.PointF vPoint = new System.Drawing.PointF(0.5F, 0.5F);
                    mPrinting.XLBarCode(vSize, vPoint, vPath);

                    //mPrinting.CountBarCodeImage = 1;
                    /////////////////////���ڵ� ����/////////////////////// /////
                     
                    vObject = pRow["ITEM_CODE"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(6, 11, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(6, 11, vConvertString);
                    }

                    //DESCRIPTION
                    vObject = pRow["ITEM_DESCRIPTION"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(7, 11, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(7, 11, vConvertString);
                    }

                    //QUANTITY
                    vObject = pRow["ONHAND_QTY"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(8, 11, "'" + vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(8, 11, vConvertString);
                    }

                    //VENDOR NAME
                    vObject = pRow["CUSTOMER_DESC"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(9, 11, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(9, 11, vConvertString);
                    }

                    //MANUFACTURED DATE
                    vObject = pRow["WEEK_DATE"];
                    IsConvert = IsConvertDate(vObject, out vDateTime);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}-{1:D2}-{2:D2}", vDateTime.Year, vDateTime.Month, vDateTime.Day);
                        mPrinting.XLSetCell(10, 11, "'" + vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Format("{0}-{1:D2}-{2:D2}", vDateTime.Year, vDateTime.Month, vDateTime.Day);
                        mPrinting.XLSetCell(10, 11, "'" + vConvertString);
                    }
                }

                else
                {
                    mPrinting.XLActiveSheet("LABEL");
                     
                    //Desc
                    vObject = pRow["ITEM_DESCRIPTION"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(1, 4, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(1, 4, vConvertString);
                    }

                    //Code
                    vObject = pRow["ITEM_CODE"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(3, 4, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(3, 4, vConvertString);
                    }

                    /////////////////////���ڵ� ����////////////////////////////
                    //Create a new PDF QR barcode
                    Syncfusion.Pdf.Barcode.PdfCode128BBarcode barcode = new Syncfusion.Pdf.Barcode.PdfCode128BBarcode();

                    //barcode = Syncfusion.Pdf.Barcode.QRCodeVersion.Auto;
                    barcode.EncodeStartStopSymbols = true;
                    barcode.EnableCheckDigit = true;
                    barcode.ShowCheckDigit = false;
                    barcode.TextDisplayLocation = Syncfusion.Pdf.Barcode.TextLocation.None;
                    barcode.BarHeight = 90;
                    barcode.NarrowBarWidth = 1;
                    barcode.BarcodeToTextGapHeight = 20;

                    //Set the barcode text
                    barcode.Text = string.Format("{0}", pRow["PACKING_BOX_NO"]);

                    //Export the barcode as image
                    Image img = barcode.ToImage();

                    string vPath = System.Environment.CurrentDirectory + @"\BARCODE.jpg";
                    img.Save(vPath, System.Drawing.Imaging.ImageFormat.Png);
                                                            
                    System.Drawing.SizeF vSize = new System.Drawing.SizeF(400F, 45F);
                    System.Drawing.PointF vPoint = new System.Drawing.PointF(20F, 65F);
                    mPrinting.XLBarCode(vSize, vPoint, vPath);

                    //Lot No
                    vObject = pRow["JOB_NO"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(8, 1, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(8, 1, vConvertString);
                    }

                    //����������
                    vObject = pRow["PRINT_DATE"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(8, 8, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(8, 8, vConvertString);
                    }

                    //����
                    vObject = pRow["WEEK_NUM"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(8, 13, "'" + vConvertString);
                        mPrinting.XLSetCell(12, 3, "'" + vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(8, 13, vConvertString);
                        mPrinting.XLSetCell(12, 3, vConvertString);
                    }

                    //Qty
                    vObject = pRow["ONHAND_QTY"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(8, 17, "'" + vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(8, 17, vConvertString);
                    }

                    //Line Data
                    vObject = pRow["LINE_DATA"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                        mPrinting.XLSetCell(9, 1, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(9, 1, vConvertString);
                    }
                }
                
                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM);

                mImageIndex = mImageIndex + 2;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            //mPageNumber = mPageNumber + 1;

            return mPageNumber;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine, string V_Print_Type)
        {
            int vPrintingLineEND = mCopyLineSUM - 9; //1~62:67���� ������ ��µǴ� ���� 62 �̹Ƿ�, 6�� ���� �ȴ�
            if (vPrintingLineEND < pPrintingLine)
            {
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //ù��° ������ ����
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            bool isOpen;
            int test = mPrinting.CountBarCodeImage;
            int vCopySumPrintingLine = pCopySumPrintingLine;
            
            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            
            object vRangeSource = pPrinting.XLGetRange(1, 1, 13, 21); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet("Destination");

            //////////////////////
            int vCountBarCodeImage = mPrinting.CountBarCodeImage;
            int vIndexBarCodeImage = mPrinting.CountBarCodeImage;
            vIndexBarCodeImage = vCountBarCodeImage;
            for (int vRow = 0; vRow < vCountBarCodeImage; vRow++)
            {
                mPrinting.XLDeleteBarCode(vIndexBarCodeImage);
                vIndexBarCodeImage--;
            }
            mPrinting.CountBarCodeImage = 1;

            ////////////////////////
            object vRangeDestination = pPrinting.XLGetRange(1, 1, 13, 21); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);
            mPrinting.CountBarCodeImage++;

            mPageNumber++; //������ ��ȣ
            string vPageNumberText = string.Format("Page    {0} of  {1}", mPageNumber, mPageTotalNumber);
            int vRowSTART = vCopyPrintingRowEnd -1;
            int vRowEND = vCopyPrintingRowEnd -1 ;
            int vColumnSTART = 30;
            int vColumnEND = 33;
            //mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
            //mPrinting.XLSetCell(vRowSTART-39 , vColumnSTART+34, vPageNumberText); //������ ��ȣ, XLcell[��, ��]

            Printing(1, 1);

            mPrinting.XLOpenFileClose();
            isOpen = XLFileOpen();

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLActiveSheet("Destination");

            mPrinting.XLPrinting(pPageSTART, pPageEND);

            
        }

        #endregion;

        #region ----- Save Methods ----

        //public void SAVE(string pSaveFileName)
        //{
        //    System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

        //    int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
        //    vMaxNumber = vMaxNumber + 1;
        //    string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

        //    vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
        //    mPrinting.XLSave(vSaveFileName);
        //}

        public void SAVE(string pSaveFileName)
        {
            try
            {
                mPrinting.XLSave(pSaveFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
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

        public void DeleteSheet(string V_PRINT_TYPE)
        {
            bool isSuccess = false;

            try
            {
                isSuccess = mPrinting.XLDeleteSheet("SourceTab_KR");
                isSuccess = mPrinting.XLDeleteSheet("SourceTab_EN");
                isSuccess = mPrinting.XLDeleteSheet("SourceTab_CH");
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Process Methods ----

        public void KillProcess_Excel()
        {
            try
            {
                string vTitleName = string.Empty;

                System.Diagnostics.Process[] vProcessFXEStart = System.Diagnostics.Process.GetProcessesByName("Excel");

                int vCountProcess = vProcessFXEStart.Length;
                if (vCountProcess > 0)
                {
                    Dispose();

                    vProcessFXEStart = System.Diagnostics.Process.GetProcessesByName("Excel");
                    vCountProcess = vProcessFXEStart.Length;
                    if (vCountProcess > 0)
                    {
                        for (int vRow = 0; vRow < vCountProcess; vRow++)
                        {
                            vTitleName = vProcessFXEStart[vRow].MainWindowTitle;
                            if (vTitleName == "")
                            {
                                vProcessFXEStart[vRow].Kill();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                mMessageError = ex.Message;
            }
        }

        #endregion;
    }
}