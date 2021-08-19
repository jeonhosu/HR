using System;
using System.Drawing;
using System.Text;
using System.Globalization;
using ISCommonUtil;

namespace PPMF0638
{
    public class PF4i_0110

    {
        #region ----- Variables -----
                
        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();        


        private System.Drawing.Printing.PrintDocument mPrintDoc = null;
        private System.Windows.Forms.PrintDialog mPrintDialog = null;
        private System.Windows.Forms.PrintPreviewDialog mPrintPreviewDialog;

        private System.Data.DataRow mRow = null;

        private string mMessageError = string.Empty;

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }        
        #endregion;

        #region ----- Constructor -----

        public PF4i_0110(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, System.Windows.Forms.PrintDialog pPrintDialog, System.Windows.Forms.PrintPreviewDialog pPrintPreviewDialog)
        {
            mAppInterface = pAppInterface;

            mPrintDialog = pPrintDialog;
            mPrintPreviewDialog = pPrintPreviewDialog;
            mPrintDoc = new System.Drawing.Printing.PrintDocument();
            /////////////////////////////////////////////////////////
            //string InBox = Convert.ToString(mRow["PACKING_BOX_NO"]);
            //string OutBox = Convert.ToString(mRow["OUT_BOX_NO"]);

            //if (InBox == OutBox)
            //{
            //    if (mPrintDoc.DefaultPageSettings.Landscape)
            //    {
            //        mPrintDoc.DefaultPageSettings.Landscape = false;

            //    }
            //}
            //else
            //{
            //    if (!mPrintDoc.DefaultPageSettings.Landscape)
            //    {
            //        mPrintDoc.DefaultPageSettings.Landscape = true;
            //    }
            //}
            /////////////////////////////////////////////////////////
            mPrintDoc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocument_PrintPage);
        }

        #endregion;

        #region ----- Print Page Event -----

        private void PrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //인쇄할 페이지가 더 있는지 체크(기본값 false)
            e.HasMorePages = false;

            string InBox = Convert.ToString(mRow["PACKING_BOX_NO"]);
            string OutBox = Convert.ToString(mRow["OUT_BOX_NO"]);
            //실제 인쇄하는 부분//
            if (InBox == OutBox)
            {
                PrintDocumentOutBox(mRow, e);
            }
            else
            {
                PrintDocumentInBox2(mRow, e);
            }

            //int endRow = mDTL.RowCount -1  ;
            //if (pageStartRow <= endRow)
            //{
            //    e.HasMorePages = true ;
            //    PrintDocument(pageStartRow, e);

            //    if (pageStartRow == endRow)
            //    {
            //        e.HasMorePages = false;
            //        pageStartRow = 0;
            //    }
            //    else
            //    {
            //        pageStartRow = pageStartRow + 1;
            //    }
            //}
        }

        #endregion;

        #region ----- Dispose Method -----

        public void Dispose()
        {
            mPrintDialog.Dispose();            
            mPrintDoc.Dispose();
        }

        #endregion;

        #region ----- PRINTING Method -----

        public void PRINTING(System.Data.DataRow pRow)
        {
            mRow = pRow;

            try
            {
                int vLable_Copies = 1;//iConv.ISNumtoZero(pRow["LABEL_PRT_COPIES"], 1);
                mPrintDialog.PrinterSettings.Copies = (short)vLable_Copies;
                
                //System.Windows.Forms.DialogResult vResult = mPrintDialog.ShowDialog();
                System.Windows.Forms.DialogResult vResult = System.Windows.Forms.DialogResult.OK;

                //short vInput_Copies = mPrintDialog.PrinterSettings.DefaultPageSettings.PrinterSettings.Copies;
                //mPrintDialog.PrinterSettings.Copies = mPrintDialog.PrinterSettings.DefaultPageSettings.PrinterSettings.Copies;
                string vPrintName = mPrintDialog.PrinterSettings.PrinterName;
                short vCopies = mPrintDialog.PrinterSettings.Copies;

                if (vResult == System.Windows.Forms.DialogResult.OK)
                {
                    mAppInterface.OnAppMessageEvent(vPrintName);
                    System.Windows.Forms.Application.DoEvents();
                      
                    mPrintDoc.PrinterSettings.PrinterName = vPrintName; //선택한 프린터 기종
                    mPrintDoc.PrinterSettings.Copies = vCopies;         //인쇄매수
                    /////////////////////////////////////////////////////////
                    //string InBox = Convert.ToString(mRow["PACKING_BOX_NO"]);
                    //string OutBox = Convert.ToString(mRow["OUT_BOX_NO"]);

                    //if (InBox == OutBox)
                    //{
                    //    if (mPrintDoc.DefaultPageSettings.Landscape)
                    //    {
                    //        mPrintDoc.DefaultPageSettings.Landscape = false;

                    //    }
                    //}
                    //else
                    //{
                    //    if (!mPrintDoc.DefaultPageSettings.Landscape)
                    //    {
                    //        mPrintDoc.DefaultPageSettings.Landscape = true;
                    //    }
                    //}
                    /////////////////////////////////////////////////////////
                    //mPrintDoc.DefaultPageSettings.PaperSize.Height = 4;
                    //mPrintDoc.DefaultPageSettings.PaperSize.Width = 10; 
                    mPrintDoc.Print();

                    
                    
                    ////인쇄 미리보기 
                    //mPrintPreviewDialog.ClientSize = new System.Drawing.Size(500, 350);
                    //mPrintPreviewDialog.PrintPreviewControl.Zoom = 120.0F / 100.0F;
                    //mPrintPreviewDialog.Document = mPrintDoc;
                    
                    //mPrintPreviewDialog.ShowDialog();
                }                
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }            
        }
        
        #endregion;

        #region ----- PRINTING Method ----

        private void printMeth(System.Drawing.Printing.PrintPageEventArgs e, string vTextPrint,int ixx,ref int iy, System.Drawing.Font vPrintFont, System.Drawing.Font vPrintFont_Small)
        {
 
            byte[] btstr = Encoding.Default.GetBytes(vTextPrint);


            if (btstr.Length > 38)
            {
                int i = Str2Line(vTextPrint, 38).Length;
                string strPrint = Str2Line(vTextPrint, 38) + "\r\n" + vTextPrint.Substring(i, vTextPrint.Length - i);
                e.Graphics.DrawString(strPrint, vPrintFont, System.Drawing.Brushes.Black, ixx, iy);
            }
            else
            {
                e.Graphics.DrawString(vTextPrint, vPrintFont, System.Drawing.Brushes.Black, ixx, iy);

            } 
        }

        //2줄 인쇄 
        public string Str2Line(string thestr, int cutlen)
        {
            string returnstr = "";
            int charcode = 0, charlen = 0;
            foreach (char c in thestr)
            {
                charcode = (int)c;
                if (charcode > 128)
                {
                    charlen = charlen + 2;
                }
                else
                {
                    charlen++;
                }
                if (charlen > cutlen)
                {
                    break;
                }
                else
                {
                    returnstr = returnstr + c.ToString();
                }
            }
            return returnstr;
        }

        private string DateToString(string str)
        {
           
            try
            {
                return "";
            }
            catch
            {
                return "";
            }
        }

        private string DateToAZ()
        {
            int idateCnt = DateTime.Now.DayOfYear;
            int AzCnt = idateCnt % 26;
            if (AzCnt == 0)
            {
                AzCnt = AzCnt + 26;
            }
            AzCnt = AzCnt + 64;
            char[] intchar = new char[1];
            intchar[0] = (char)AzCnt;
            string rtnStr = new string(intchar);

            return rtnStr;
        }

        private static int GetWeekCnt()
        {
            DateTime dt = DateTime.Now;
            int week = Enum.GetValues(typeof(DayOfWeek)).Length;
            int dayOffset = (int)dt.AddDays(-(dt.Day - 1)).DayOfWeek;
            int weekCnt = (dt.Day + dayOffset) / week;
            weekCnt += ((dt.Day + dayOffset) % week) > 0 ? 1 : 0;
            return weekCnt;
        }

        public static int WeeksInYear()
        {
            DateTime date = DateTime.Now;
            GregorianCalendar cal = new GregorianCalendar(GregorianCalendarTypes.Localized);
            return cal.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }


        private void PrintDocumentOutBox(System.Data.DataRow pRow, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //LINE 긋기 위한 위치 지정 
            float vLine_Default_X = 15F;//= (float)(iConv.ISDecimaltoZero(pRow["START_X"]));    // X축 START 위치 지정 //
            float vLine_Default_Y = 20F;//= (float)(iConv.ISDecimaltoZero(pRow["START_Y"]));    // Y축 START 위치 지정 //            

            float vLine_Max_X = 380F;            
            float vLine_Max_Y = 235F;

            float vLine_X = vLine_Default_X;    // X축 START 위치 지정 //
            
            //float vLIne_X_1st = vLine_Default_X + 80F;
            float vLIne_X_1st = vLine_Default_X + 67F;
            float vLIne_X_2nd = vLine_Default_X + 208F;
            float vLIne_X_3rd = vLine_Default_X + 268.5F;
            
            float vLine_Default_Y_Add = 17F;
            float vLine_Y = vLine_Default_Y;    // Y축 START 위치 지정 //

            float vLine_X_End = vLine_Max_X + vLine_Default_X;                      // X축 너비에 X축 시작포인트를 합함
            float vLine_Y_End = vLine_Max_Y + vLine_Default_Y;                      // Y축 너비에 Y축 시작포인트를 합함

            float vFont_X_Add = 3F;
            float vFont_Y_Add = 4;
            float vFont_X = 0;
            float vFont_Y = 0;

            string vTextPrint = string.Empty;

            //숫자 오른쪽 정렬 위치 지정
            float x = 183.0F;
            float y = 70.0F;
            float width = 200.0F;
            float height = 50.0F;
            RectangleF drawRect = new RectangleF(x, y, width, height);

            RectangleF autoNextLine_Item = new RectangleF(85, 22, 300, 31);
            RectangleF autoNextLine_Spec = new RectangleF(85, 55, 300, 15);

            StringFormat drawFormat = new StringFormat();
            drawFormat.Alignment = StringAlignment.Far;

            try
            {
                //System.Drawing.Font vBarCodeFont = new Font("Free 3 of 9 Extended", 23F, FontStyle.Regular, GraphicsUnit.Pixel);
                System.Drawing.Font vPrintFont_Title = new Font("굴림", 18F, FontStyle.Bold, GraphicsUnit.Point, ((Byte)(129)));
                System.Drawing.Font vPrintFont_Month = new Font("굴림", 40F, FontStyle.Bold, GraphicsUnit.Point, ((Byte)(129)));
                System.Drawing.Font vPrintFont_Bold = new Font("굴림", 8F, FontStyle.Bold, GraphicsUnit.Point, ((Byte)(129)));
                System.Drawing.Font vPrintFont = new System.Drawing.Font("굴림", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_Bar = new System.Drawing.Font("굴림", 6F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_Small = new System.Drawing.Font("굴림", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_PackingBoxNo = new System.Drawing.Font("바탕체", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_Item = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));

                //------------------------------------//
                // 0. 외곽 LINE - 가로 맨위, 맨밑 선긋기
                //------------------------------------//
                Drawing_Rectangle(e, new Pen(Color.Black, 1), vLine_X, vLine_Y, vLine_Max_X, vLine_Max_Y);


                //------------------------------------//
                // 1. 상단 MATERIAL TAG  (RoHS)
                //------------------------------------//
                vFont_X = vLine_X + vFont_X_Add;
                vFont_Y = vLine_Y + vFont_Y_Add;
                vTextPrint = "MATERIAL TAG   (RoHS)";
                e.Graphics.DrawString(vTextPrint, vPrintFont_Title, System.Drawing.Brushes.Black, vFont_X, vFont_Y);




















                
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        //private void PrintDocumentInBox(System.Data.DataRow pRow, System.Drawing.Printing.PrintPageEventArgs e)
        //{
        //    //LINE 긋기 위한 위치 지정 
        //    float vLine_Default_X = 35F;//= (float)(iConv.ISDecimaltoZero(pRow["START_X"]));    // X축 START 위치 지정 //
        //    float vLine_Default_Y = 10F;//= (float)(iConv.ISDecimaltoZero(pRow["START_Y"]));    // Y축 START 위치 지정 //            

        //    float vLine_Max_X = 155F;
        //    float vLine_Max_Y = 195F;

        //    float vLine_X_Bottom = 5F;
        //    float vLine_Y_Bottom = 212F;
        //    float vLine_Max_X_Bottom = 155F;
        //    float vLine_Max_Y_Bottom = 194.5F;
        //    float vLine_Square_Space = 1.54F;

        //    float vLine_X = vLine_Default_X;    // X축 START 위치 지정 //

        //    //float vLIne_X_1st = vLine_Default_X + 57F;
        //    float vLIne_X_1st = vLine_Default_X + 45F;
        //    float vLIne_X_2nd = vLine_Default_X + 208F;
        //    float vLIne_X_3rd = vLine_Default_X + 268.5F;

        //    float vLine_Default_Y_Add = 17F;
        //    float vLine_Y = vLine_Default_Y;    // Y축 START 위치 지정 //

        //    float vLine_X_End = vLine_Max_X + vLine_Default_X;                      // X축 너비에 X축 시작포인트를 합함
        //    float vLine_Y_End = vLine_Max_Y + vLine_Default_Y;                      // Y축 너비에 Y축 시작포인트를 합함

        //    float vFont_X_Add = 3F;
        //    float vFont_Y_Add = 4F;
        //    float vFont_X = 0;
        //    float vFont_Y = 0;

        //    float vIn_Line_Width = 17.72F + vLine_Default_Y;
        //    float vIn_Line_Width_Add = 17.72F;
        //    float vIn_Line_Height = 17.72F;
        //    float vIn_Line_Height_Add = 17.72F;
        //    float vIn_Font_Height_Space = 1F;
        //    float vIn_Font_Left_Space = 4F;
        //    float vIn_Line_Height_2 = 0F;

        //    string vTextPrint = string.Empty;

        //    //숫자 오른쪽 정렬 위치 지정
        //    RectangleF drawRect_Qty = new RectangleF(35, 162, 63, 300);
        //    RectangleF autoNextLine_Desc = new RectangleF(84, 25, 110, 100);
        //    RectangleF autoNextLine_Lot = new RectangleF(84, 115, 112, 300);
        //    RectangleF drawRect_Qty_2 = new RectangleF(85, 252, 63, 300);
        //    RectangleF autoNextLine_Desc_2 = new RectangleF(84, 293, 110, 368);
        //    RectangleF autoNextLine_Lot_2 = new RectangleF(84, 383, 112, 568);

        //    StringFormat drawFormat = new StringFormat();
        //    drawFormat.Alignment = StringAlignment.Far;

        //    try
        //    {
        //        //System.Drawing.Font vBarCodeFont = new Font("Free 3 of 9 Extended", 23F, FontStyle.Regular, GraphicsUnit.Pixel);
        //        System.Drawing.Font vPrintFont_Bold = new Font("굴림", 8F, FontStyle.Bold, GraphicsUnit.Point, ((Byte)(129)));
        //        System.Drawing.Font vPrintFont = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
        //        System.Drawing.Font vPrintFont_Bar = new System.Drawing.Font("굴림", 6F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
        //        System.Drawing.Font vPrintFont_Small = new System.Drawing.Font("굴림", 5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
        //        System.Drawing.Font vPrintFont_PackingBoxNo = new System.Drawing.Font("바탕체", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
        //        System.Drawing.Font vPrintFont_Month = new System.Drawing.Font("굴림", 38F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));

        //        //------------------------------------//
        //        //FIRST BOX
        //        Drawing_Rectangle(e, new Pen(Color.Black, 1), vLine_X, vLine_Y, vLine_Max_X, vLine_Max_Y);

        //        //1.1 Title 인쇄 
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 6;
        //        vFont_Y = vLine_Y + vFont_Y_Add + 11;
        //        vTextPrint = "Item";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //1.2 선(l) 인쇄                
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLIne_X_1st, vLine_Y, vLIne_X_1st, vLine_Y + vIn_Line_Height);

        //        //1.3 DATA 인쇄
        //        vTextPrint = iConv.ISNull(pRow["ITEM_DESCRIPTION"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, autoNextLine_Desc);

        //        //1.4 선(ㅡ) 인쇄
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End, vIn_Line_Width);

        //        //2.1 Title 인쇄 
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 1;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = "Item No";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //2.2 선(l) 인쇄                
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLIne_X_1st, vLine_Y, vLIne_X_1st, vLine_Y + vIn_Line_Height);

        //        //2.3 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space - vIn_Line_Height_Add;
        //        vTextPrint = iConv.ISNull(pRow["ITEM_CODE"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //2.4 선(ㅡ) 인쇄
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End, vIn_Line_Width);

        //        //3.1 DATA 인쇄                
        //        vFont_X = 40;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        string vPACKAGING_BOX_NO = iConv.ISNull(pRow["PACKING_BOX_NO"]);

        //        //3.2 이미지를 이용한 BarCode Code128(바코드 인쇄)
        //        MakeBarCode(e, vPACKAGING_BOX_NO, (int)vFont_X, (int)vFont_Y, 146, 20);

        //        //3.3 DATA 인쇄
        //        vFont_X = vLIne_X_1st + vFont_X_Add - 2;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space + 21;
        //        vTextPrint = iConv.ISNull(pRow["PACKING_BOX_NO"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_PackingBoxNo, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //3.4 선(ㅡ) 인쇄
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        //Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End, vIn_Line_Width);
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End, vIn_Line_Width);

        //        //4.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + (vIn_Line_Height_Add * 2);
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 2;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + 11;
        //        vTextPrint = "Lot No";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //4.2 DATA 인쇄
        //        vTextPrint = iConv.ISNull(pRow["MAT_JOB_NO"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, autoNextLine_Lot);

        //        //5.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + (vIn_Line_Height_Add * 2);
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = "Supplier";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //5.2 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = iConv.ISNull(pRow["VENDOR_FULL_NAME"]);
        //        if (vTextPrint.Length > 26)
        //        {
        //            vTextPrint = vTextPrint.Substring(0, 26);
        //        }
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //6.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 8;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = "Qty";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //6.2 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = string.Format("{0:#,###,##0}", pRow["PACKING_QTY"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, drawRect_Qty, drawFormat);

        //        //7.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space - 1;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = "Input Date";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //7.2 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = iConv.ISNull(pRow["DELIVERY_DATE"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //8.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 6;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = "Spec";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //8.2 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = iConv.ISNull(pRow["ITEM_SPECIFICATION"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //9.1 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add + 38;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space - 40;
        //        vTextPrint = iConv.ISNull(pRow["DELIVERY_MONTH"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Month, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //10.1 선(l) 인쇄 Lot No ~ 규격               
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLIne_X_1st, 104F, vLIne_X_1st, 210F);

        //        //10.2 선(l) 인쇄 수량 ~ 규격 중간선               
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLIne_X_1st + 50, 157F, vLIne_X_1st + 50, 210F);


        //        //10.3 선(ㅡ) 인쇄 Lot No ~ 규격
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        //Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End, vIn_Line_Width);//LOT_NO
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End, vIn_Line_Width);//LOT_NO
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End, vIn_Line_Width);//공급업체
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End - 60, vIn_Line_Width);//수량
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End - 60, vIn_Line_Width);//입고일자
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width, vLine_X_End, vIn_Line_Width);//규격

        //        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //        //SECOND BOX
        //        Drawing_Rectangle(e, new Pen(Color.Black, 1), vLine_X_Bottom, vLine_Y_Bottom, vLine_Max_X_Bottom, vLine_Max_Y_Bottom);

        //        //11.1 선(ㅡ) 인쇄
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End - 60, vIn_Line_Width + vLine_Square_Space);//규격                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End - 60, vIn_Line_Width + vLine_Square_Space);//입고일자                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//수량                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//공급업체                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        //Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//품명                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//품명                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//자재코드                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        //Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//바코드                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//바코드                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        //Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//LOT NO                
        //        vIn_Line_Width = vIn_Line_Width + vIn_Line_Width_Add;
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vIn_Line_Width + vLine_Square_Space, vLine_X_End, vIn_Line_Width + vLine_Square_Space);//LOT NO                

        //        //11.2 선(l) 인쇄                                
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLIne_X_1st, 212F, vLIne_X_1st, 335F);
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLIne_X_1st, 371F, vLIne_X_1st, 407F);
        //        Drawing_Line(e, new Pen(Color.Black, 1), vLIne_X_1st + 50, 212F, vLIne_X_1st + 50, 265F);

        //        //12.1 Title 인쇄 
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 6;
        //        vFont_Y = vLine_Y + vFont_Y_Add + 198;
        //        vTextPrint = "Spec";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //12.2 DATA 인쇄 
        //        vIn_Line_Height_2 = vIn_Line_Height - vIn_Line_Height_Add;
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height_2 + vIn_Font_Height_Space + vLine_Square_Space;
        //        vTextPrint = iConv.ISNull(pRow["ITEM_SPECIFICATION"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //13.1 Title 인쇄 
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space - 1;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space + vLine_Square_Space;
        //        vTextPrint = "Input Date";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //13.2 DATA 인쇄 
        //        vIn_Line_Height_2 = vIn_Line_Height_2 + vIn_Line_Height_Add;
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height_2 + vIn_Font_Height_Space + vLine_Square_Space;
        //        vTextPrint = iConv.ISNull(pRow["DELIVERY_DATE"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //14.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 8;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space + vLine_Square_Space;
        //        vTextPrint = "Qty";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //14.2 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space;
        //        vTextPrint = string.Format("{0:#,###,##0}", pRow["PACKING_QTY"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, drawRect_Qty_2, drawFormat);

        //        //15.1 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add + 38;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space - 38;
        //        vTextPrint = iConv.ISNull(pRow["DELIVERY_MONTH"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Month, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //16.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space + vLine_Square_Space;
        //        vTextPrint = "Supplier";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //16.2 DATA 인쇄 
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space + vLine_Square_Space;
        //        vTextPrint = iConv.ISNull(pRow["VENDOR_FULL_NAME"]);
        //        if (vTextPrint.Length > 26)
        //        {
        //            vTextPrint = vTextPrint.Substring(0, 26);
        //        }
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //17.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 6;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space + vLine_Square_Space + 10;
        //        vTextPrint = "Item";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //17.2 DATA 인쇄               
        //        vTextPrint = iConv.ISNull(pRow["ITEM_DESCRIPTION"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, autoNextLine_Desc_2);

        //        //18.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 1;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space + vLine_Square_Space;
        //        vTextPrint = "Item No";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //18.2 DATA 인쇄 
        //        vIn_Line_Height_2 = vIn_Line_Height_2 + vIn_Line_Height_Add;
        //        vIn_Line_Height_2 = vIn_Line_Height_2 + vIn_Line_Height_Add;
        //        vIn_Line_Height_2 = vIn_Line_Height_2 + vIn_Line_Height_Add;
        //        vIn_Line_Height_2 = vIn_Line_Height_2 + vIn_Line_Height_Add;
        //        vIn_Line_Height_2 = vIn_Line_Height_2 + vIn_Line_Height_Add;
        //        vFont_X = vLIne_X_1st + vFont_X_Add;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height_2 + vIn_Font_Height_Space + vLine_Square_Space;
        //        vTextPrint = iConv.ISNull(pRow["ITEM_CODE"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //19.1 DATA 인쇄                
        //        vIn_Line_Height_2 = vIn_Line_Height_2 + vIn_Line_Height_Add;
        //        vFont_X = 10;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height_2 + vIn_Font_Height_Space + vLine_Square_Space;
        //        vPACKAGING_BOX_NO = iConv.ISNull(pRow["PACKING_BOX_NO"]);

        //        //19.2 이미지를 이용한 BarCode Code128(바코드 인쇄)
        //        MakeBarCode(e, vPACKAGING_BOX_NO, (int)vFont_X, (int)vFont_Y, 146, 20);

        //        //19.3 DATA 인쇄
        //        vFont_X = vLIne_X_1st + vFont_X_Add - 2;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height_2 + vIn_Font_Height_Space + 22;
        //        vTextPrint = iConv.ISNull(pRow["PACKING_BOX_NO"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_PackingBoxNo, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //20.1 Title 인쇄 
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vIn_Line_Height = vIn_Line_Height + vIn_Line_Height_Add;
        //        vFont_X = vLine_X + vFont_X_Add + vIn_Font_Left_Space + 6;
        //        vFont_Y = vLine_Y + vFont_Y_Add + vIn_Line_Height + vIn_Font_Height_Space + vLine_Square_Space + 10;
        //        vTextPrint = "Lot No";
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

        //        //20.2 DATA 인쇄
        //        vTextPrint = iConv.ISNull(pRow["MAT_JOB_NO"]);
        //        e.Graphics.DrawString(vTextPrint, vPrintFont_Small, System.Drawing.Brushes.Black, autoNextLine_Lot_2);
        //    }
        //    catch (System.Exception ex)
        //    {
        //        mMessageError = ex.Message;
        //        mAppInterface.OnAppMessageEvent(mMessageError);
        //        System.Windows.Forms.Application.DoEvents();
        //    }
        //}

        private void PrintDocumentInBox2(System.Data.DataRow pRow, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //LINE 긋기 위한 위치 지정 
            float vLine_Default_X = 15F;//= (float)(iConv.ISDecimaltoZero(pRow["START_X"]));    // X축 START 위치 지정 //
            float vLine_Default_Y = 20F;//= (float)(iConv.ISDecimaltoZero(pRow["START_Y"]));    // Y축 START 위치 지정 //            

            float vLine_Max_X = 450F;
            float vLine_Max_Y = 450F;

            float vLine_X = vLine_Default_X;    // X축 START 위치 지정 //

            float vLIne_X_1st = vLine_Default_X + 45F;
            float vLIne_X_2nd = vLine_Default_X + 208F;
            float vLIne_X_3rd = vLine_Default_X + 268.5F;
            float vLIne_X_4rd = vLine_Default_X + 310F;

            float vLIne_X_Column = 90;

            float vLine_Default_Y_Add = 17F;
            float vLine_Y = vLine_Default_Y;    // Y축 START 위치 지정 //

            float vLine_X_End = vLine_Max_X + vLine_Default_X;                      // X축 너비에 X축 시작포인트를 합함
            float vLine_Y_End = vLine_Max_Y + vLine_Default_Y;                      // Y축 너비에 Y축 시작포인트를 합함

            float vFont_X_Add = 3F;
            float vFont_Y_Add = 4;
            float vL_Font_X_Add = 100F;
            float vL_Font_Y_Add = 4;
            float vFont_X = 0;
            float vFont_Y = 0;

            float vSecond_X = (float)190.5;

            //    //숫자 오른쪽 정렬 위치 지정
            //    RectangleF drawRect_Qty = new RectangleF(35, 162, 63, 300);
            //    RectangleF autoNextLine_Desc = new RectangleF(84, 25, 110, 100);
            //    RectangleF autoNextLine_Lot = new RectangleF(84, 115, 112, 300);
            //    RectangleF drawRect_Qty_2 = new RectangleF(85, 252, 63, 300);
            //    RectangleF autoNextLine_Desc_2 = new RectangleF(84, 293, 110, 368);
            //    RectangleF autoNextLine_Lot_2 = new RectangleF(84, 383, 112, 568);

            string vTextPrint = string.Empty;

            //숫자 오른쪽 정렬 위치 지정
            float x = 183.0F;
            float y = 70.0F;
            float width = 200.0F;
            float height = 50.0F;
            RectangleF drawRect = new RectangleF(x, y, width, height);
            RectangleF autoNextLine_Item = new RectangleF(119, 90, 270, 31);
            //RectangleF autoNextLine_Item_1 = new RectangleF(245, 123, 150, 31);
            RectangleF autoNextLine_Lot = new RectangleF(55, 90, 270, 15);
            //RectangleF autoNextLine_Lot_1 = new RectangleF(245, 107, 150, 15);

            StringFormat drawFormat = new StringFormat();
            drawFormat.Alignment = StringAlignment.Far;

            //이미지 정의
            //Image image_bsk_logo = Image.FromFile(@"C:\Infosolution\FLEX_ERP_BSK\PROD\Image\BskLogo.png");
            Graphics grfx = e.Graphics;

            try
            {
                //System.Drawing.Font vBarCodeFont = new Font("Free 3 of 9 Extended", 23F, FontStyle.Regular, GraphicsUnit.Pixel);
                System.Drawing.Font vPrintFont_Bold = new Font("굴림", 5F, FontStyle.Bold, GraphicsUnit.Point, ((Byte)(129)));
                System.Drawing.Font vPrintFont_Title = new Font("굴림", 13F, FontStyle.Bold, GraphicsUnit.Point, ((Byte)(129)));
                System.Drawing.Font vPrintFont_Line = new Font("굴림", 12F, FontStyle.Bold, GraphicsUnit.Point, ((Byte)(129))); //10
                System.Drawing.Font vPrintFont = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_Bar = new System.Drawing.Font("굴림", 6F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_Small = new System.Drawing.Font("굴림", 5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_Large = new System.Drawing.Font("굴림", 50F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_PackingBoxNo = new System.Drawing.Font("바탕체", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_Month = new System.Drawing.Font("굴림", 30F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                System.Drawing.Font vPrintFont_Item = new Font("굴림", 7F, FontStyle.Bold, GraphicsUnit.Point, ((Byte)(129)));
                System.Drawing.Font vPrintFont_Date = new System.Drawing.Font("굴림", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
                //------------------------------------//
                // 0. 외곽 LINE - 가로 맨위, 맨밑 선긋기
                //------------------------------------//
                //Drawing_Rectangle(e, new Pen(Color.Black, 1), vLine_X, vLine_Y, vLine_Max_X, vLine_Max_Y);

                //------------------------------------//
                //1. 박스 Number
                //------------------------------------//
                vFont_X = vLine_X + 10;
                vFont_Y = vLine_Y + vL_Font_Y_Add;
                vTextPrint = iConv.ISNull(pRow["ITEM_DESCRIPTION"]);
                e.Graphics.DrawString(vTextPrint, vPrintFont_Line, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

                vLine_Y = vLine_Y + vLine_Default_Y_Add;

                // 1. 제품명
                vFont_X = vLine_X + 10;
                vFont_Y = vLine_Y + vL_Font_Y_Add + 10;
                vTextPrint = "DESC  : " + iConv.ISNull(pRow["ITEM_DESCRIPTION"]);
                e.Graphics.DrawString(vTextPrint, vPrintFont_Line, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

                vLine_Y = vLine_Y + 60;

                // 2. lot 번호
                vFont_X = vLine_X + 10;
                vFont_Y = vLine_Y + vL_Font_Y_Add + 20;
                vTextPrint = "Lot No : " + iConv.ISNull(pRow["JOB_NO"]);
                e.Graphics.DrawString(vTextPrint, vPrintFont_Line, System.Drawing.Brushes.Black, vFont_X, vLine_Y);

                // 선(--) 긋기
                vLine_Y = vLine_Y + 10;
             
                //3. Barcode
                vFont_X = vLine_X + 33;//vLIne_X_1st + vFont_X_Add;
                vFont_Y = vLine_Y + vL_Font_Y_Add + 10;//vLine_Y + vFont_Y_Add;
                string vPACKAGING_BOX_NO = iConv.ISNull(pRow["PACKING_BOX_NO"]);
                MakeBarCode(e, vPACKAGING_BOX_NO, 50, 120, 340, 30);

                vLine_Y = vLine_Y + 30;
                //e.Graphics.DrawString(vPACKAGING_BOX_NO, vPrintFont_Line, System.Drawing.Brushes.Black, vFont_X, vLine_Y);


                // 4. 포장일
                vFont_X = vLine_X + 10;
                vFont_Y = vLine_Y + vL_Font_Y_Add + 12;
                vTextPrint = iConv.ISNull(pRow["DELIVERY_DATE"]);
                if (vTextPrint != "")
                {
                    System.DateTime vDateTime = Convert.ToDateTime(vTextPrint);
                    vTextPrint = "DATE   : " + vDateTime.ToString("yyyy-MM-dd", null);
                }
                
                e.Graphics.DrawString(vTextPrint, vPrintFont_Line, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

                vLine_Y = vLine_Y + 16;

                // 5.수량
                vFont_X = vLine_X + 10;
                vFont_Y = vLine_Y + vL_Font_Y_Add + 12;
                vTextPrint = "QTY    : " + iConv.ISNull(pRow["ONHAND_QTY"]);
                e.Graphics.DrawString(vTextPrint, vPrintFont_Line, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

                // 5.3 선(--) 긋기
                vLine_Y = vLine_Y + 20;
                //Drawing_Line(e, new Pen(Color.Black, 1), vLine_X, vLine_Y + vL_Font_Y_Add + 4, vLine_Max_X + 15, vLine_Y + vL_Font_Y_Add + 4);

                // 6. 제조회사
                vFont_X = vLine_X + 10;
                vFont_Y = vLine_Y + vL_Font_Y_Add + 15;
                vTextPrint = "SIFLEX CO., LTD";
                e.Graphics.DrawString(vTextPrint, vPrintFont_Line, System.Drawing.Brushes.Black, vFont_X, vFont_Y);

                // 6.1 선(l) 긋기
                //Drawing_Line(e, new Pen(Color.Black, 1), vLIne_X_Column, vLine_Y + vL_Font_Y_Add + 4, vLIne_X_Column, vLine_Y + vL_Font_Y_Add + 37);


          
              

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Draw 라인 -----

        private void Drawing_Line(System.Drawing.Printing.PrintPageEventArgs e, Pen pPen, float pX1, float pY1, float pX2, float pY2)
        {
            e.Graphics.DrawLine(pPen, pX1, pY1, pX2, pY2);
        }
        #endregion;

        #region ----- Draw 사각형 -----

        private void Drawing_Rectangle(System.Drawing.Printing.PrintPageEventArgs e, Pen pPen, float pX1, float pY1, float pWidth, float pHeight)
        {
            e.Graphics.DrawRectangle(pPen, pX1, pY1, pWidth, pHeight);
        }
        #endregion;

        #region ----- Draw 채우기 사각형 -----

        private void Drawing_FillRectangle(System.Drawing.Printing.PrintPageEventArgs e, Brush pBrush, float pX1, float pY1, float pWidth, float pHeight)
        {
            e.Graphics.FillRectangle(pBrush, pX1, pY1, pWidth, pHeight);
        }
        #endregion;

        #region ----- Draw BarCode Method -----

        private bool MakeBarCode(System.Drawing.Printing.PrintPageEventArgs e, string pPACKAGING_LABEL_NO, int pFR_X, int pFR_Y, int pTO_X, int pTO_Y)
        {
            bool IsDraw = false;

            BarCode128.Code128Encode vBarCode = new BarCode128.Code128Encode();

            int vBarWeight = 1;            
            string vEncodingString = pPACKAGING_LABEL_NO;   //바코드 생성 대상 

            try
            {
                vBarCode.WeightBar = vBarWeight;
                //vBarCode.IsBottomText = true;                   //바코드 하단 바코드 내용 표시 여부 
                vBarCode.IsBottomText = false;                   
                vBarCode.ColorBottomText = System.Drawing.Brushes.Black;
                System.Drawing.Bitmap vImageBarCode = vBarCode.EncodingBarCode(vEncodingString);

                System.Drawing.Rectangle vRect = new Rectangle(pFR_X, pFR_Y, pTO_X, pTO_Y);
                e.Graphics.DrawImage(vImageBarCode, vRect);

                vImageBarCode.Dispose();

                IsDraw = true;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
                return IsDraw;
            }
            return IsDraw;
        }

        #endregion;

        #region ----- Make BarCode Method : 이하 코드는 사용 안함 ----

        private bool MakeBarCode(string pPath, string pBarCodeName, string pString)
        {
            bool isMake = false;

            BarCode128.Code128Encode vBarCode = new BarCode128.Code128Encode();

            try
            {
                int vBarWeight = 1;
                string vEncodingString = pString;

                vBarCode.WeightBar = vBarWeight;
                vBarCode.IsBottomText = true;
                vBarCode.ColorBottomText = System.Drawing.Brushes.DarkBlue;
                System.Drawing.Bitmap vImageBarCode = vBarCode.EncodingBarCode(vEncodingString);
                
                
                
                if (vImageBarCode != null)
                {
                    string vSaveFileName = pBarCodeName;
                    bool isSvae = vBarCode.Save(vImageBarCode, pPath, vSaveFileName);
                    if (isSvae != true)
                    {
                        mAppInterface.OnAppMessageEvent("Save false");
                        System.Windows.Forms.Application.DoEvents();
                    }
                    else
                    {
                        isMake = true;
                    }

                    vImageBarCode.Dispose();
                }
                else
                {
                    mAppInterface.OnAppMessageEvent("BarCode Make Error");
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return isMake;
        }

        private Bitmap BarCode(string pPath, string pBarCodeName, string pString)
        {
            BarCode128.Code128Encode vBarCode = new BarCode128.Code128Encode();

            try
            {
                int vBarWeight = 1;
                string vEncodingString = pString;

                vBarCode.WeightBar = vBarWeight;
                vBarCode.IsBottomText = true;
                vBarCode.ColorBottomText = System.Drawing.Brushes.DarkBlue;
                System.Drawing.Bitmap vImageBarCode = vBarCode.EncodingBarCode(vEncodingString);

                return vImageBarCode;


                //if (vImageBarCode != null)
                //{
                //    string vSaveFileName = pBarCodeName;
                //    bool isSvae = vBarCode.Save(vImageBarCode, pPath, vSaveFileName);
                //    if (isSvae != true)
                //    {
                //        mAppInterface.OnAppMessageEvent("Save false");
                //        System.Windows.Forms.Application.DoEvents();
                //    }
                //    else
                //    {
                //        isMake = true;
                //    }

                //    vImageBarCode.Dispose();
                //}
                //else
                //{
                //    mAppInterface.OnAppMessageEvent("BarCode Make Error");
                //    System.Windows.Forms.Application.DoEvents();
                //}
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return null;
        }

        #endregion;
    }
}
