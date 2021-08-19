using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using Syncfusion.Pdf.Barcode;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;

namespace ZZF001
{
    public partial class ZZF001 : Office2007Form
    {
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISCommonUtil.ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISCommonUtil.ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public ZZF001()
        {
            InitializeComponent();
        }

        public ZZF001(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (TB_MAIN.SelectedTab.TabIndex == TP_ITEM_MASTER.TabIndex)
            {
                IDA_ITEM_MASTER.Fill();
                ISG_ITEM_MASTER.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_ITEM_RECEIPT.TabIndex)
            {
                IDA_ITEM_RECEIPT.Fill();
                ISG_ITEM_RECEIPT.Focus();
            }
        }

        private void Insert_DB()
        {
            ISG_ITEM_MASTER.SetCellValue("ENABLED_FLAG", "Y");
            ISG_ITEM_MASTER.SetCellValue("EFFECTIVE_DATE_FR",  iDate.ISMonth_1st(DateTime.Today));
        }

        private void Insert_Receipt()
        {
            ISG_ITEM_MASTER.SetCellValue("RECEIPT_QTY", 0);
            ISG_ITEM_MASTER.SetCellValue("RECEIPT_DATE", iDate.ISMonth_1st(DateTime.Today));
        }

        private void Init_Amount(object pUnit_Price, object pRecipt_Qty)
        {
            IDC_GET_RECEIPT_AMT_P.SetCommandParamValue("W_UNIT_PRICE", pUnit_Price);
            IDC_GET_RECEIPT_AMT_P.SetCommandParamValue("W_RECEIPT_QTY", pRecipt_Qty);
            IDC_GET_RECEIPT_AMT_P.ExecuteNonQuery();
            object vRECEIPT_AMT = IDC_GET_RECEIPT_AMT_P.GetCommandParamValue("O_RECEIPT_AMT");
            ISG_ITEM_RECEIPT.SetCellValue("RECEIPT_AMT", vRECEIPT_AMT);
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_ITEM_MASTER.IsFocused)
                    {
                        IDA_ITEM_MASTER.AddOver();
                        Insert_DB();
                    }
                    else if(IDA_ITEM_RECEIPT.IsFocused)
                    {
                        IDA_ITEM_RECEIPT.AddOver();
                        Insert_Receipt();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_ITEM_MASTER.IsFocused)
                    {
                        IDA_ITEM_MASTER.AddUnder();
                        Insert_DB();
                    }
                    else if(IDA_ITEM_RECEIPT.IsFocused)
                    {
                        IDA_ITEM_RECEIPT.AddUnder();
                        Insert_Receipt();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_ITEM_MASTER.Update();
                    IDA_ITEM_RECEIPT.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_ITEM_MASTER.IsFocused)
                    {
                        IDA_ITEM_MASTER.Cancel();
                    }
                    else if(IDA_ITEM_RECEIPT.IsFocused)
                    {
                        IDA_ITEM_RECEIPT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_ITEM_MASTER.IsFocused)
                    {
                        IDA_ITEM_MASTER.Delete();
                    }
                    else if(IDA_ITEM_RECEIPT.IsFocused)
                    {
                        IDA_ITEM_RECEIPT.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ---- Form Event ----    
        
        private void ZZF001_Load(object sender, EventArgs e)
        {
            W_RECEIPT_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            W_RECEIPT_DATE_TO.EditValue = DateTime.Today;

            IDA_ITEM_MASTER.FillSchema();
            IDA_ITEM_RECEIPT.FillSchema();
        }

        private void ISG_ITEM_RECEIPT_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == ISG_ITEM_RECEIPT.GetColumnToIndex("RECEIPT_QTY"))
            {
                Init_Amount(ISG_ITEM_RECEIPT.GetCellValue("UNIT_PRICE"), e.NewValue); 
            }
        }
        #endregion;

        #region ---- Lookup Event ----- 

        private void ILA_ITEM_MASTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ITEM_MASTER.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ILA_ITEM_MASTER_W2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ITEM_MASTER.SetLookupParamValue("W_ENABLED_FLAG", "N");
        }

        private void ILA_ITEM_MASTER_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ITEM_MASTER.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_ITEM_MASTER_2_SelectedRowData(object pSender)
        {
            Init_Amount(ISG_ITEM_RECEIPT.GetCellValue("UNIT_PRICE"), ISG_ITEM_RECEIPT.GetCellValue("RECEIPT_QTY"));
        }

        #endregion

        private void isButton1_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //Create a new PDF QR barcode
            Syncfusion.Pdf.Barcode.PdfQRBarcode barcode = new Syncfusion.Pdf.Barcode.PdfQRBarcode();

            barcode.Version = Syncfusion.Pdf.Barcode.QRCodeVersion.Auto;
            barcode.ErrorCorrectionLevel = Syncfusion.Pdf.Barcode.PdfErrorCorrectionLevel.High;
            
            barcode.XDimension = 5f;
            //Set the barcode text
            barcode.Text = iConv.ISNull(isEditAdv1.EditValue);
            //Export the barcode as image
            Image img = barcode.ToImage();// (Image)new Bitmap(barcode.ToImage(), new Size(450, 450));

            string vPath = System.Environment.CurrentDirectory + @"\..\..\QR.jpg";

            //Save the image to stream
            img.Save(vPath, System.Drawing.Imaging.ImageFormat.Png);
            pictureBox1.ImageLocation = vPath;
        }

        private void isButton2_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //Create a new PDF barcode
            Syncfusion.Pdf.Barcode.PdfCode128BBarcode barcode = new Syncfusion.Pdf.Barcode.PdfCode128BBarcode();
             
            //barcode = Syncfusion.Pdf.Barcode.QRCodeVersion.Auto;
            barcode.EncodeStartStopSymbols = true;
            barcode.EnableCheckDigit = true;
            barcode.ShowCheckDigit = false;
            barcode.BarHeight = 90;
            barcode.NarrowBarWidth = 1;
            barcode.BarcodeToTextGapHeight = 20;


            //Set the barcode text
            barcode.Text = string.Format("{0}", iConv.ISNull(isEditAdv2.EditValue));
            
            //Export the barcode as image
            Image img = barcode.ToImage();// (Image)new Bitmap(barcode.ToImage(), new Size(450, 450));

            string vPath = System.Environment.CurrentDirectory + @"\..\..\BARCODE.jpg";

            //Save the image to stream
            img.Save(vPath, System.Drawing.Imaging.ImageFormat.Png);
            pictureBox2.ImageLocation = vPath;

        }

        private void isButton3_ButtonClick(object pSender, EventArgs pEventArgs)
        {
             
            //Create a new PDF document.
            PdfDocument document = new PdfDocument();
            
            //Creates a new page and adds it as the last page of the document
            PdfPage page = document.Pages.Add();
            
            //Creates a new PDF datamatrix barcode.
            PdfDataMatrixBarcode datamatrix = new PdfDataMatrixBarcode();
            
            //Sets the barcode text.
            datamatrix.Text = string.Format("{0}", iConv.ISNull(isEditAdv3.EditValue));

            //Set the dimention of the barcode.
            datamatrix.XDimension = 5;
            
            //Set barcode size.
            datamatrix.Size = PdfDataMatrixSize.Size20x20;

            ////Set the barcode location.
            //datamatrix.Location = new PointF(100, 100);
            
            ////Draws a barcode on the new Page.
            //datamatrix.Draw(page);
            datamatrix.Draw(page, new PointF(10, 30));

            //Export the barcode as image
            //Image img = datamatrix.ToImage();// (Image)new Bitmap(barcode.ToImage(), new Size(450, 450));

            //string vPath = System.Environment.CurrentDirectory + @"\..\..\BARCODE.jpg";
            //img.Save(vPath, System.Drawing.Imaging.ImageFormat.Png);
            //pictureBox2.ImageLocation = vPath; 

            string vPath = System.Environment.CurrentDirectory + @"\..\..\BARCODE.jpg";
            //Save the PDF document
            document.Save(vPath);

            //Close the document.
            document.Close(true);
        }
    }
}