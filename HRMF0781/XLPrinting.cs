using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace HRMF0781
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
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "Source1";
        private string mSourceSheet2 = "Source2";

        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // ù ������ üũ.

        // �μ�� ���ο� �հ�.
        private int mCopyLineSUM = 0;

        // �μ� - ��ȭ �μ� ����.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 31;
        private int mCopy_EndRow = 53;
        private int mPrintingLastRow = 52;  //���� ������ �μ� ���� ����.

        private int mCurrentRow = 12;        //���� �μ�Ǵ� row ��ġ.
        private int mDefaultPageRow = 11;    //������ skip�� ����Ǵ� �⺻ PageCount �⺻��.

        // �μ�2 - �ҵ漼 ���μ� �μ� ����.
        private int mCopy_StartCol2 = 1;
        private int mCopy_StartRow2 = 1;
        private int mCopy_EndCol2 = 33;
        private int mCopy_EndRow2 = 59;
        private int mPrintingLastRow2 = 59;  //���� ������ �μ� ���� ����.

        private int mCurrentRow2 = 12;        //���� �μ�Ǵ� row ��ġ.
        private int mDefaultPageRow2 = 11;    //������ skip�� ����Ǵ� �⺻ PageCount �⺻��.

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

        #region ----- Array Set 0 ----

        private void SetArray0(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// �׸����� �÷��� ���� �÷��ε��� �� ����
            pGDColumn = new int[3];
            pXLColumn = new int[3];
            // �׸��� or �ƴ��� ��ġ.
            pGDColumn[0] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[1] = pGrid.GetColumnToIndex("GL_AMOUNT");
            pGDColumn[2] = pGrid.GetColumnToIndex("VAT_AMOUNT");

            // ������ �μ��ؾ� �� ��ġ.
            pXLColumn[0] = 12;
            pXLColumn[1] = 22;
            pXLColumn[2] = 34;
        }

        #endregion;

        #region ----- Array Set 1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// �׸����� �÷��� ���� �÷��ε��� �� ����
            pGDColumn = new int[12];
            pXLColumn = new int[12];
            // �׸��� or �ƴ��� ��ġ.
            pGDColumn[0] = pGrid.GetColumnToIndex("PERSON_NUM");
            pGDColumn[1] = pGrid.GetColumnToIndex("NAME");
            pGDColumn[2] = pGrid.GetColumnToIndex("REPRE_NUM");
            pGDColumn[3] = pGrid.GetColumnToIndex("DEPT_NAME");
            pGDColumn[4] = pGrid.GetColumnToIndex("FLOOR_NAME");
            pGDColumn[5] = pGrid.GetColumnToIndex("ABIL_NAME");
            pGDColumn[6] = pGrid.GetColumnToIndex("POST_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("ORI_JOIN_DATE");
            pGDColumn[8] = pGrid.GetColumnToIndex("JOIN_DATE");
            pGDColumn[9] = pGrid.GetColumnToIndex("RETIRE_DATE");
            pGDColumn[10] = pGrid.GetColumnToIndex("CONTINUE_YEAR");
            pGDColumn[11] = pGrid.GetColumnToIndex("END_SCH_NAME");


            // ������ �μ��ؾ� �� ��ġ.
            pXLColumn[0] = 1;
            pXLColumn[1] = 6;
            pXLColumn[2] = 11;
            pXLColumn[3] = 17;
            pXLColumn[4] = 24;
            pXLColumn[5] = 31;
            pXLColumn[6] = 36;
            pXLColumn[7] = 42;
            pXLColumn[8] = 46;
            pXLColumn[9] = 50;
            pXLColumn[10] = 54;
            pXLColumn[11] = 59;
        }

        #endregion;

        #region ----- Array Set 2  : Adapter ����� ----

        //private void SetArray2(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        //{// �ƴ����� table ��.
        //    pGDColumn = new int[10];
        //    pXLColumn = new int[10];

        //    pGDColumn[0] = pTable.Columns.IndexOf("PO_TYPE_NAME");
        //    pGDColumn[1] = pTable.Columns.IndexOf("DISPLAY_NAME");
        //    pGDColumn[2] = pTable.Columns.IndexOf("PO_DATE");
        //    pGDColumn[3] = pTable.Columns.IndexOf("PO_NO");
        //    pGDColumn[4] = pTable.Columns.IndexOf("SUPPLIER_SHORT_NAME");
        //    pGDColumn[5] = pTable.Columns.IndexOf("PRICE_TERM_NAME");
        //    pGDColumn[6] = pTable.Columns.IndexOf("PAYMENT_METHOD_NAME");
        //    pGDColumn[7] = pTable.Columns.IndexOf("PAYMENT_TERM_NAME");
        //    pGDColumn[8] = pTable.Columns.IndexOf("REMARK");
        //    pGDColumn[9] = pTable.Columns.IndexOf("STEP_DESCRIPTION");


        //    pXLColumn[0] = 9;   //PO_TYPE_NAME
        //    pXLColumn[1] = 25;  //DISPLAY_NAME
        //    pXLColumn[2] = 42;  //PO_DATE
        //    pXLColumn[3] = 54;  //PO_NO
        //    pXLColumn[4] = 9;   //SUPPLIER_SHORT_NAME
        //    pXLColumn[5] = 35;  //PRICE_TERM_NAME
        //    pXLColumn[6] = 14;  //PAYMENT_METHOD_NAME
        //    pXLColumn[7] = 41;  //PAYMENT_TERM_NAME
        //    pXLColumn[8] = 7;   //REMARK
        //    pXLColumn[9] = 49;  //�ݾ�
        //}

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

        #region ----- Excel Write -----

        #region ----- Header Write Method ----

        public void HeaderWrite(System.Data.DataRow pRow)
        {// ��� �μ�.
            int vXLine = 0;
            int vXLColumn = 0;
            object vValue = null;
            string vString = string.Empty;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);
                //�ͼӳ⵵
                vXLine = 3;
                vXLColumn = 4;
                vValue = pRow["WITHHOLDING_YEAR"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //���ܱ���-������
                vXLine = 3;
                vXLColumn = 40;
                vValue = pRow["NATIONALITY_1"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //���ܱ���-�ܱ���
                vXLine = 4;
                vXLColumn = 40;
                vValue = pRow["NATIONALITY_9"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //��������
                vXLine = 5;
                vXLColumn = 31;
                vValue = pRow["NATION_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //������CODE
                vXLine = 5;
                vXLColumn = 38;
                vValue = pRow["NATION_ISO_CODE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //¡���ǹ��� ����ڵ�Ϲ�ȣ
                vXLine = 8;
                vXLColumn = 12;
                vValue = pRow["VAT_NUMBER"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //¡���ǹ��� ���θ�
                vXLine = 8;
                vXLColumn = 26;
                vValue = pRow["CORP_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //¡���ǹ��� ��ǥ�ڸ�
                vXLine = 8;
                vXLColumn = 38;
                vValue = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //¡���ǹ��� �ֹ�(����)��Ϲ�ȣ
                vXLine = 9;
                vXLColumn = 12;
                vValue = pRow["LEGAL_NUMBER"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //¡���ǹ��� �ּ�
                vXLine = 9;
                vXLColumn = 26;
                vValue = pRow["CORP_ADDRESS"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //�ҵ��� ��ȣ
                vXLine = 10;
                vXLColumn = 12;
                vValue = pRow["COMPANY_NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //�ҵ��� ����ڵ�Ϲ�ȣ
                vXLine = 10;
                vXLColumn = 33;
                vValue = pRow["TAX_REG_NO"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //�ҵ��� ����� �ּ�
                vXLine = 11;
                vXLColumn = 12;
                vValue = pRow["COMPANY_ADDRESS"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //�ҵ��� ����
                vXLine = 12;
                vXLColumn = 12;
                vValue = pRow["NAME"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //�ҵ��� �ֹε�Ϲ�ȣ
                vXLine = 12;
                vXLColumn = 33;
                vValue = pRow["REPRE_NUM"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //�ҵ��� �ּ�
                vXLine = 13;
                vXLColumn = 12;
                vValue = pRow["ADDRESS"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //��������
                vXLine = 14;
                vXLColumn = 7;
                vValue = pRow["BUSINESS_CODE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //�μ�����
                vXLine = 30;
                vXLColumn = 17;
                vValue = pRow["PRINT_DATE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //¡���ǹ���
                vXLine = 31;
                vXLColumn = 21;
                vValue = pRow["WITHHOLDING_AGENT"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

                //������
                vXLine = 32;
                vXLColumn = 1;
                vValue = pRow["TAX_OFFICE"];
                if (iString.ISNull(vValue) != string.Empty)
                {
                    vString = string.Format("{0}", vValue);
                }
                else
                {
                    vString = null;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vString);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Header1 (�հ�) Write Method ----

        private void XLHeader1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int[] pGDColumn, int[] pXLColumn)
        {// ��� �μ�.
            int vXLine = 0; //������ ������ ǥ�õǴ� �� ��ȣ

            int vIDX_VAT_TYPE = pGrid.GetColumnToIndex("VAT_TYPE");
            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            // ���Ǵ� ���� ����.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            { // ������ �����ؼ� Ÿ�� �� ������ ����.(
                mPrinting.XLActiveSheet(mTargetSheet);

                for (int i = 0; i < pGrid.RowCount; i++)
                {
                    // ���հ� ���п� ���� �μ� ROW ����.
                    if ("T" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//���հ�
                        vXLine = 9;
                    }
                    else if ("3" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//�ſ�ī��.
                        vXLine = 13;
                    }
                    else if ("11" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//���ݿ�����.
                        vXLine = 10;
                    }

                    //0 - �ŷ��Ǽ�.
                    vGDColumnIndex = pGDColumn[0];
                    vXLColumnIndex = pXLColumn[0];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //1 - ���ް���
                    vGDColumnIndex = pGDColumn[1];
                    vXLColumnIndex = pXLColumn[1];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    //2 - ����
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

        #region ----- Excel Write [KRW] Method -----

        private int LineWrite(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //�Ű���.
                //�ſ�
                vXLine = 3;
                vXLColumn = 2;
                vObject = pRow["MONTHLY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }
                
                //�ݱ�
                vXLine = 3;
                vXLColumn = 5;
                vObject = pRow["HALF_YEARLY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //����
                vXLine = 3;
                vXLColumn = 6;
                vObject = pRow["MODIFY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //����
                vXLine = 3;
                vXLColumn = 7;
                vObject = pRow["YEAR_END_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //�ҵ�ó��
                vXLine = 3;
                vXLColumn = 9;
                vObject = pRow["INCOME_DISPOSED_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //ȯ�޽�û
                vXLine = 3;
                vXLColumn = 10;
                vObject = pRow["REFUND_REQUEST_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }


                //�ͼӿ���
                vXLine = 2;
                vXLColumn = 28;
                vObject = pRow["STD_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���޿���
                vXLine = 3;
                vXLColumn = 28;
                vObject = pRow["PAY_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���θ�
                vXLine = 4;
                vXLColumn = 8;
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //��ǥ��
                vXLine = 4;
                vXLColumn = 17;
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ϰ����ο���
                vXLine = 4;
                vXLColumn = 28;
                vObject = pRow["ALL_PAYMENT_YN"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ڴ�����������
                vXLine = 5;
                vXLColumn = 28;
                vObject = pRow["BUSINESS_UNIT_TAX_YN"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ڵ�Ϲ�ȣ
                vXLine = 6;
                vXLColumn = 8;
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����� �ּ�
                vXLine = 6;
                vXLColumn = 17;
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ȭ��ȣ
                vXLine = 6;
                vXLColumn = 27;
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�̸��� �ּ�
                vXLine = 7;
                vXLColumn = 27;
                vObject = pRow["EMAIL"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //���̼��� �ο���
                vXLine = 12;
                vXLColumn = 10;
                vObject = pRow["A01_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���̼��� �����޾�
                vXLine = 12;
                vXLColumn = 12;
                vObject = pRow["A01_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���̼��� �ҵ漼��
                vXLine = 12;
                vXLColumn = 16;
                vObject = pRow["A01_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));
                                
                //���̼��� ��Ư��
                vXLine = 12;
                vXLColumn = 20;
                vObject = pRow["A01_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���̼��� ���꼼
                vXLine = 12;
                vXLColumn = 23;
                vObject = pRow["A01_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� �ο���
                vXLine = 13;
                vXLColumn = 10;
                vObject = pRow["A02_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� �����޾�
                vXLine = 13;
                vXLColumn = 12;
                vObject = pRow["A02_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� �ҵ漼��
                vXLine = 13;
                vXLColumn = 16;
                vObject = pRow["A02_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� ��Ư��
                vXLine = 13;
                vXLColumn = 20;
                vObject = pRow["A02_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� ���꼼
                vXLine = 13;
                vXLColumn = 23;
                vObject = pRow["A02_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�Ͽ�ٷ� �ο���
                vXLine = 14;
                vXLColumn = 10;
                vObject = pRow["A03_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�Ͽ�ٷ� �����޾�
                vXLine = 14;
                vXLColumn = 12;
                vObject = pRow["A03_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�Ͽ�ٷ� �ҵ漼��
                vXLine = 14;
                vXLColumn = 16;
                vObject = pRow["A03_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�Ͽ�ٷ� ���꼼
                vXLine = 14;
                vXLColumn = 23;
                vObject = pRow["A03_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �ο���
                vXLine = 15;
                vXLColumn = 10;
                vObject = pRow["A04_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �����޾�
                vXLine = 15;
                vXLColumn = 12;
                vObject = pRow["A04_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �ҵ漼��
                vXLine = 15;
                vXLColumn = 16;
                vObject = pRow["A04_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ��Ư��
                vXLine = 15;
                vXLColumn = 20;
                vObject = pRow["A04_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���꼼
                vXLine = 15;
                vXLColumn = 23;
                vObject = pRow["A04_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ �ο���
                vXLine = 16;
                vXLColumn = 10;
                vObject = pRow["A10_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ �����޾�
                vXLine = 16;
                vXLColumn = 12;
                vObject = pRow["A10_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ �ҵ漼��
                vXLine = 16;
                vXLColumn = 16;
                vObject = pRow["A10_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ��Ư��
                vXLine = 16;
                vXLColumn = 20;
                vObject = pRow["A10_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ���꼼
                vXLine = 16;
                vXLColumn = 23;
                vObject = pRow["A10_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ��� ���� ȯ�޼���
                vXLine = 16;
                vXLColumn = 25;
                vObject = pRow["A10_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 16;
                vXLColumn = 27;
                vObject = pRow["A10_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ���μ��� ��Ư��
                vXLine = 16;
                vXLColumn = 30;
                vObject = pRow["A10_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //----------------------------------------------------------------------------------
                //�����ҵ� ���ݰ��� �ο� 
                vXLine = 17;
                vXLColumn = 10;
                vObject = pRow["A21_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���ݰ��� �����޾�
                vXLine = 17;
                vXLColumn = 12;
                vObject = pRow["A21_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���ݰ��� �ҵ漼��
                vXLine = 17;
                vXLColumn = 16;
                vObject = pRow["A21_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���ݰ��� ���꼼
                vXLine = 17;
                vXLColumn = 23;
                vObject = pRow["A21_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���ݰ��� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 17;
                vXLColumn = 27;
                vObject = pRow["A21_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� �ο� 
                vXLine = 18;
                vXLColumn = 10;
                vObject = pRow["A22_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� �����޾�
                vXLine = 18;
                vXLColumn = 12;
                vObject = pRow["A22_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� �ҵ漼��
                vXLine = 18;
                vXLColumn = 16;
                vObject = pRow["A22_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� ���꼼
                vXLine = 18;
                vXLColumn = 23;
                vObject = pRow["A22_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 18;
                vXLColumn = 27;
                vObject = pRow["A22_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //�����ҵ� �ο���
                vXLine = 19;
                vXLColumn = 10;
                vObject = pRow["A20_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �����޾�
                vXLine = 19;
                vXLColumn = 12;
                vObject = pRow["A20_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �ҵ漼��
                vXLine = 19;
                vXLColumn = 16;
                vObject = pRow["A20_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���꼼
                vXLine = 19;
                vXLColumn = 23;
                vObject = pRow["A20_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ��� ���� ȯ�޼���
                vXLine = 19;
                vXLColumn = 25;
                vObject = pRow["A20_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 19;
                vXLColumn = 27;
                vObject = pRow["A20_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                
                //------------------------------------------------------------------------------------------
                //����ҵ� �ſ�¡�� �ο���
                vXLine = 20;
                vXLColumn = 10;
                vObject = pRow["A25_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �ſ�¡�� �����޾�
                vXLine = 20;
                vXLColumn = 12;
                vObject = pRow["A25_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �ſ�¡�� �ҵ漼��
                vXLine = 20;
                vXLColumn = 16;
                vObject = pRow["A25_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �ſ�¡�� ���꼼
                vXLine = 20;
                vXLColumn = 23;
                vObject = pRow["A25_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //����ҵ� �������� �ο���
                vXLine = 21;
                vXLColumn = 10;
                vObject = pRow["A26_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �������� �����޾�
                vXLine = 21;
                vXLColumn = 12;
                vObject = pRow["A26_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �������� �ҵ漼��
                vXLine = 21;
                vXLColumn = 16;
                vObject = pRow["A26_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �������� ��Ư��
                vXLine = 21;
                vXLColumn = 20;
                vObject = pRow["A26_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �������� ���꼼
                vXLine = 21;
                vXLColumn = 23;
                vObject = pRow["A26_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //����ҵ� ������ �ο���
                vXLine = 22;
                vXLColumn = 10;
                vObject = pRow["A30_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ �����޾�
                vXLine = 22;
                vXLColumn = 12;
                vObject = pRow["A30_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ �ҵ漼��
                vXLine = 22;
                vXLColumn = 16;
                vObject = pRow["A30_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ��Ư��
                vXLine = 22;
                vXLColumn = 20;
                vObject = pRow["A30_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ���꼼
                vXLine = 22;
                vXLColumn = 23;
                vObject = pRow["A30_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ���� ȯ�޼���
                vXLine = 22;
                vXLColumn = 25;
                vObject = pRow["A30_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 22;
                vXLColumn = 27;
                vObject = pRow["A30_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ���μ��� ��Ư��
                vXLine = 22;
                vXLColumn = 30;
                vObject = pRow["A30_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //------------------------------------------------------------------------------------------
                //��Ÿ�ҵ� ���ݰ��� �ο���
                vXLine = 23;
                vXLColumn = 10;
                vObject = pRow["A41_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ���ݰ��� �����޾�
                vXLine = 23;
                vXLColumn = 12;
                vObject = pRow["A41_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ���ݰ��� �ҵ漼��
                vXLine = 23;
                vXLColumn = 16;
                vObject = pRow["A41_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ���ݰ��� ���꼼
                vXLine = 23;
                vXLColumn = 23;
                vObject = pRow["A41_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ���ݰ��� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 23;
                vXLColumn = 27;
                vObject = pRow["A41_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� �ο���
                vXLine = 24;
                vXLColumn = 10;
                vObject = pRow["A42_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� �����޾�
                vXLine = 24;
                vXLColumn = 12;
                vObject = pRow["A42_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� �ҵ漼��
                vXLine = 24;
                vXLColumn = 16;
                vObject = pRow["A42_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� ���꼼
                vXLine = 24;
                vXLColumn = 23;
                vObject = pRow["A42_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 24;
                vXLColumn = 27;
                vObject = pRow["A42_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));
                
                //��Ÿ�ҵ� ������ �ο���
                vXLine = 25;
                vXLColumn = 10;
                vObject = pRow["A40_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ �����޾�
                vXLine = 25;
                vXLColumn = 12;
                vObject = pRow["A40_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ �ҵ漼��
                vXLine = 25;
                vXLColumn = 16;
                vObject = pRow["A40_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ ���꼼
                vXLine = 25;
                vXLColumn = 23;
                vObject = pRow["A40_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ ���� ȯ�޼���
                vXLine = 25;
                vXLColumn = 25;
                vObject = pRow["A40_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 25;
                vXLColumn = 27;
                vObject = pRow["A40_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //------------------------------------------------------------------------------------------
                //���ݼҵ� ���ݰ��� �ο���
                vXLine = 26;
                vXLColumn = 10;
                vObject = pRow["A48_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ���ݰ��� �����޾�
                vXLine = 26;
                vXLColumn = 12;
                vObject = pRow["A48_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ���ݰ��� �ҵ漼��
                vXLine = 26;
                vXLColumn = 16;
                vObject = pRow["A48_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ���ݰ��� ���꼼
                vXLine = 26;
                vXLColumn = 23;
                vObject = pRow["A48_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ��������(�ſ�) �ο���
                vXLine = 27;
                vXLColumn = 10;
                vObject = pRow["A45_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ��������(�ſ�) �����޾�
                vXLine = 27;
                vXLColumn = 12;
                vObject = pRow["A45_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ��������(�ſ�) �ҵ漼��
                vXLine = 27;
                vXLColumn = 16;
                vObject = pRow["A45_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ��������(�ſ�) ���꼼
                vXLine = 27;
                vXLColumn = 23;
                vObject = pRow["A45_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //���ݼҵ� �������� �ο���
                vXLine = 28;
                vXLColumn = 10;
                vObject = pRow["A46_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� �������� �����޾�
                vXLine = 28;
                vXLColumn = 12;
                vObject = pRow["A46_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� �������� �ҵ漼��
                vXLine = 28;
                vXLColumn = 16;
                vObject = pRow["A46_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� �������� ���꼼
                vXLine = 28;
                vXLColumn = 23;
                vObject = pRow["A46_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));
                
                //���ݼҵ� ������ �ο���
                vXLine = 29;
                vXLColumn = 10;
                vObject = pRow["A47_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ �����޾�
                vXLine = 29;
                vXLColumn = 12;
                vObject = pRow["A47_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ �ҵ漼��
                vXLine = 29;
                vXLColumn = 16;
                vObject = pRow["A47_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ ���꼼
                vXLine = 29;
                vXLColumn = 23;
                vObject = pRow["A47_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ ���� ȯ�޼���
                vXLine = 29;
                vXLColumn = 25;
                vObject = pRow["A47_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 29;
                vXLColumn = 27;
                vObject = pRow["A47_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //----------------------------------------------------------------------------
                //���ڼҵ� �ο���
                vXLine = 30;
                vXLColumn = 10;
                vObject = pRow["A50_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� �����޾�
                vXLine = 30;
                vXLColumn = 12;
                vObject = pRow["A50_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� �ҵ漼��
                vXLine = 30;
                vXLColumn = 16;
                vObject = pRow["A50_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ��Ư��
                vXLine = 30;
                vXLColumn = 20;
                vObject = pRow["A50_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));
                
                //���ڼҵ� ���꼼
                vXLine = 30;
                vXLColumn = 23;
                vObject = pRow["A50_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ���� ȯ�޼���
                vXLine = 30;
                vXLColumn = 25;
                vObject = pRow["A50_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 30;
                vXLColumn = 27;
                vObject = pRow["A50_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ���μ��� ��Ư��
                vXLine = 30;
                vXLColumn = 30;
                vObject = pRow["A50_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //-------------------------------------------------------------------------------
                //���ҵ� �ο���
                vXLine = 31;
                vXLColumn = 10;
                vObject = pRow["A60_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� �����޾�
                vXLine = 31;
                vXLColumn = 12;
                vObject = pRow["A60_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� �ҵ漼��
                vXLine = 31;
                vXLColumn = 16;
                vObject = pRow["A60_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ��Ư��
                vXLine = 31;
                vXLColumn = 20;
                vObject = pRow["A60_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ���꼼
                vXLine = 31;
                vXLColumn = 23;
                vObject = pRow["A60_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ���� ȯ�޼���
                vXLine = 31;
                vXLColumn = 25;
                vObject = pRow["A60_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 31;
                vXLColumn = 27;
                vObject = pRow["A60_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ���μ��� ��Ư��
                vXLine = 31;
                vXLColumn = 30;
                vObject = pRow["A60_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //----------------------------------------------------------------------------------
                //�������� �ο���
                vXLine = 32;
                vXLColumn = 10;
                vObject = pRow["A69_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �ҵ漼��
                vXLine = 32;
                vXLColumn = 16;
                vObject = pRow["A69_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���꼼
                vXLine = 32;
                vXLColumn = 23;
                vObject = pRow["A69_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���� ȯ�޼���
                vXLine = 32;
                vXLColumn = 25;
                vObject = pRow["A69_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 32;
                vXLColumn = 27;
                vObject = pRow["A69_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //-------------------------------------------------------------------------------
                //������� �ο���
                vXLine = 33;
                vXLColumn = 10;
                vObject = pRow["A70_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� �����޾�
                vXLine = 33;
                vXLColumn = 12;
                vObject = pRow["A70_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� �ҵ漼��
                vXLine = 33;
                vXLColumn = 16;
                vObject = pRow["A70_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� ���꼼
                vXLine = 33;
                vXLColumn = 23;
                vObject = pRow["A70_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� ���� ȯ�޼���
                vXLine = 33;
                vXLColumn = 25;
                vObject = pRow["A70_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 33;
                vXLColumn = 27;
                vObject = pRow["A70_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //-------------------------------------------------------------------------------
                //���ܱ��ι��ο�õ �ο���
                vXLine = 34;
                vXLColumn = 10;
                vObject = pRow["A80_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ �����޾�
                vXLine = 34;
                vXLColumn = 12;
                vObject = pRow["A80_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ �ҵ漼��
                vXLine = 34;
                vXLColumn = 16;
                vObject = pRow["A80_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ ���꼼
                vXLine = 34;
                vXLColumn = 23;
                vObject = pRow["A80_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ ���� ȯ�޼���
                vXLine = 34;
                vXLColumn = 25;
                vObject = pRow["A80_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 34;
                vXLColumn = 27;
                vObject = pRow["A80_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //---------------------------------------------------------------------
                //�����Ű� �ҵ漼��
                vXLine = 35;
                vXLColumn = 16;
                vObject = pRow["A90_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ��Ư��
                vXLine = 35;
                vXLColumn = 20;
                vObject = pRow["A90_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ���꼼
                vXLine = 35;
                vXLColumn = 23;
                vObject = pRow["A90_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ���� ȯ�޼���
                vXLine = 35;
                vXLColumn = 25;
                vObject = pRow["A90_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 35;
                vXLColumn = 27;
                vObject = pRow["A90_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ���μ��� ��Ư��
                vXLine = 35;
                vXLColumn = 30;
                vObject = pRow["A90_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //---------------------------------------------------------------------------------
                //���հ� �ο���
                vXLine = 36;
                vXLColumn = 10;
                vObject = pRow["A99_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� �����޾�
                vXLine = 36;
                vXLColumn = 12;
                vObject = pRow["A99_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� �ҵ漼��
                vXLine = 36;
                vXLColumn = 16;
                vObject = pRow["A99_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ��Ư��
                vXLine = 36;
                vXLColumn = 20;
                vObject = pRow["A99_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ���꼼
                vXLine = 36;
                vXLColumn = 23;
                vObject = pRow["A99_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ���� ȯ�޼���
                vXLine = 36;
                vXLColumn = 25;
                vObject = pRow["A99_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 36;
                vXLColumn = 27;
                vObject = pRow["A99_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ���μ��� ��Ư��
                vXLine = 36;
                vXLColumn = 30;
                vObject = pRow["A99_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //--------------------------------------------------------------------------------
                //12.������ȯ�޼���
                vXLine = 41;
                vXLColumn = 2;
                vObject = pRow["RECEIVE_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //13.��ȯ�޽�û�Ѽ���
                vXLine = 41;
                vXLColumn = 6;
                vObject = pRow["ALREADY_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //14.�����ܾ�
                vXLine = 41;
                vXLColumn = 9;
                vObject = pRow["REFUND_BALANCE_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //15.�Ϲ�ȯ��
                vXLine = 41;
                vXLColumn = 12;
                vObject = pRow["GENERAL_REFUND_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //16.��Ź���(����ȸ���)
                vXLine = 41;
                vXLColumn = 15;
                vObject = pRow["FINANCIAL_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //17-1.�׹��� ȯ�޼���-����ȸ���
                vXLine = 41;
                vXLColumn = 18;
                vObject = pRow["ETC_REFUND_FINANCIAL_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //17-2.�׹��� ȯ�޼���-�պ���
                vXLine = 41;
                vXLColumn = 20;
                vObject = pRow["ETC_REFUND_MERGER_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));
                
                //18.�������ȯ�޼���
                vXLine = 41;
                vXLColumn = 22;
                vObject = pRow["ADJUST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));
                
                //19. ��� ������ȯ�޼���
                vXLine = 41;
                vXLColumn = 25;
                vObject = pRow["THIS_ADJUST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //20. ���� �̿� ȯ�޼���
                vXLine = 41;
                vXLColumn = 27;
                vObject = pRow["NEXT_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //21.ȯ�޽�û��
                vXLine = 41;
                vXLColumn = 30;
                vObject = pRow["REQUEST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //--------------------------------------------------------------
                //��������
                vXLine = 45;
                vXLColumn = 10;
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�Ű���
                vXLine = 46;
                vXLColumn = 7;
                vObject = pRow["WITHHOLDING_AGENT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������
                vXLine = 52;
                vXLColumn = 3;
                vObject = pRow["TAX_OFFICE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        private int LineWrite_11(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //Ÿ��Ʋ
                vXLine = 2;
                vXLColumn = 12;                
                if (iString.ISNull(pRow["REQUEST_REFUND_FLAG"]) == "Y")
                {
                    vObject = "�� ��õ¡�������Ȳ�Ű�\r\n�� ��õ¡������ȯ�޽�û��";
                }
                else
                {
                    vObject = "�� ��õ¡�������Ȳ�Ű�\r\n�� ��õ¡������ȯ�޽�û��";
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //�Ű���.
                //�ſ�
                vXLine = 3;
                vXLColumn = 2;
                vObject = pRow["MONTHLY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //�ݱ�
                vXLine = 3;
                vXLColumn = 5;
                vObject = pRow["HALF_YEARLY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //����
                vXLine = 3;
                vXLColumn = 6;
                vObject = pRow["MODIFY_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //����
                vXLine = 3;
                vXLColumn = 7;
                vObject = pRow["YEAR_END_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //�ҵ�ó��
                vXLine = 3;
                vXLColumn = 9;
                vObject = pRow["INCOME_DISPOSED_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }

                //ȯ�޽�û
                vXLine = 3;
                vXLColumn = 10;
                vObject = pRow["REFUND_REQUEST_YN"];
                if (iString.ISNull(vObject) == "Y")
                {
                    mPrinting.XLCellColorBrush(vXLine, vXLColumn, vXLine, vXLColumn, System.Drawing.Color.DarkGray);
                }


                //�ͼӿ���
                vXLine = 2;
                vXLColumn = 28;
                vObject = pRow["STD_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���޿���
                vXLine = 3;
                vXLColumn = 28;
                vObject = pRow["PAY_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���θ�
                vXLine = 4;
                vXLColumn = 8;
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ǥ��
                vXLine = 4;
                vXLColumn = 17;
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ϰ����ο���
                vXLine = 4;
                vXLColumn = 28;
                vObject = pRow["ALL_PAYMENT_YN"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ڴ�����������
                vXLine = 5;
                vXLColumn = 28;
                vObject = pRow["BUSINESS_UNIT_TAX_YN"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ڵ�Ϲ�ȣ
                vXLine = 6;
                vXLColumn = 8;
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����� �ּ�
                vXLine = 6;
                vXLColumn = 17;
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ȭ��ȣ
                vXLine = 6;
                vXLColumn = 27;
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�̸��� �ּ�
                vXLine = 7;
                vXLColumn = 27;
                vObject = pRow["EMAIL"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //���̼��� �ο���
                vXLine = 12;
                vXLColumn = 10;
                vObject = pRow["A01_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���̼��� �����޾�
                vXLine = 12;
                vXLColumn = 12;
                vObject = pRow["A01_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���̼��� �ҵ漼��
                vXLine = 12;
                vXLColumn = 16;
                vObject = pRow["A01_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���̼��� ��Ư��
                vXLine = 12;
                vXLColumn = 20;
                vObject = pRow["A01_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���̼��� ���꼼
                vXLine = 12;
                vXLColumn = 23;
                vObject = pRow["A01_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� �ο���
                vXLine = 13;
                vXLColumn = 10;
                vObject = pRow["A02_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� �����޾�
                vXLine = 13;
                vXLColumn = 12;
                vObject = pRow["A02_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� �ҵ漼��
                vXLine = 13;
                vXLColumn = 16;
                vObject = pRow["A02_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� ��Ư��
                vXLine = 13;
                vXLColumn = 20;
                vObject = pRow["A02_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ߵ���� ���꼼
                vXLine = 13;
                vXLColumn = 23;
                vObject = pRow["A02_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�Ͽ�ٷ� �ο���
                vXLine = 14;
                vXLColumn = 10;
                vObject = pRow["A03_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�Ͽ�ٷ� �����޾�
                vXLine = 14;
                vXLColumn = 12;
                vObject = pRow["A03_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�Ͽ�ٷ� �ҵ漼��
                vXLine = 14;
                vXLColumn = 16;
                vObject = pRow["A03_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�Ͽ�ٷ� ���꼼
                vXLine = 14;
                vXLColumn = 23;
                vObject = pRow["A03_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �հ� 
                vXLine = 15;
                vXLColumn = 10;
                vObject = pRow["A04_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �հ� �����޾�
                vXLine = 15;
                vXLColumn = 12;
                vObject = pRow["A04_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �հ� �ҵ漼��
                vXLine = 15;
                vXLColumn = 16;
                vObject = pRow["A04_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �հ� ��Ư��
                vXLine = 15;
                vXLColumn = 20;
                vObject = pRow["A04_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �հ� ���꼼
                vXLine = 15;
                vXLColumn = 23;
                vObject = pRow["A04_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �г� ��û�� �ο���
                vXLine = 16;
                vXLColumn = 10;
                vObject = pRow["A05_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �г� �����޾�
                vXLine = 16;
                vXLColumn = 12;
                vObject = pRow["A05_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �г� �ҵ漼��
                vXLine = 16;
                vXLColumn = 16;
                vObject = pRow["A05_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �г� ��Ư��
                vXLine = 16;
                vXLColumn = 20;
                vObject = pRow["A05_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �г� ���꼼
                vXLine = 16;
                vXLColumn = 23;
                vObject = pRow["A05_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���� �ο���
                vXLine = 17;
                vXLColumn = 10;
                vObject = pRow["A06_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���� �����޾�
                vXLine = 17;
                vXLColumn = 12;
                vObject = pRow["A06_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���� �ҵ漼��
                vXLine = 17;
                vXLColumn = 16;
                vObject = pRow["A06_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���� ��Ư��
                vXLine = 17;
                vXLColumn = 20;
                vObject = pRow["A06_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���� ���꼼
                vXLine = 17;
                vXLColumn = 23;
                vObject = pRow["A06_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ �ο���
                vXLine = 18;
                vXLColumn = 10;
                vObject = pRow["A10_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ �����޾�
                vXLine = 18;
                vXLColumn = 12;
                vObject = pRow["A10_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ �ҵ漼��
                vXLine = 18;
                vXLColumn = 16;
                vObject = pRow["A10_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ��Ư��
                vXLine = 18;
                vXLColumn = 20;
                vObject = pRow["A10_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ���꼼
                vXLine = 18;
                vXLColumn = 23;
                vObject = pRow["A10_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ��� ���� ȯ�޼���
                vXLine = 18;
                vXLColumn = 25;
                vObject = pRow["A10_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 18;
                vXLColumn = 27;
                vObject = pRow["A10_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�ٷμҵ� ������ ���μ��� ��Ư��
                vXLine = 18;
                vXLColumn = 30;
                vObject = pRow["A10_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //----------------------------------------------------------------------------------
                //�����ҵ� ���ݰ��� �ο� 
                vXLine = 19;
                vXLColumn = 10;
                vObject = pRow["A21_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���ݰ��� �����޾�
                vXLine = 19;
                vXLColumn = 12;
                vObject = pRow["A21_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���ݰ��� �ҵ漼��
                vXLine = 19;
                vXLColumn = 16;
                vObject = pRow["A21_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���ݰ��� ���꼼
                vXLine = 19;
                vXLColumn = 23;
                vObject = pRow["A21_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���ݰ��� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 19;
                vXLColumn = 27;
                vObject = pRow["A21_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� �ο� 
                vXLine = 20;
                vXLColumn = 10;
                vObject = pRow["A22_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� �����޾�
                vXLine = 20;
                vXLColumn = 12;
                vObject = pRow["A22_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� �ҵ漼��
                vXLine = 20;
                vXLColumn = 16;
                vObject = pRow["A22_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� ���꼼
                vXLine = 20;
                vXLColumn = 23;
                vObject = pRow["A22_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �׿� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 20;
                vXLColumn = 27;
                vObject = pRow["A22_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //�����ҵ� �ο���
                vXLine = 21;
                vXLColumn = 10;
                vObject = pRow["A20_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �����޾�
                vXLine = 21;
                vXLColumn = 12;
                vObject = pRow["A20_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� �ҵ漼��
                vXLine = 21;
                vXLColumn = 16;
                vObject = pRow["A20_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���꼼
                vXLine = 21;
                vXLColumn = 23;
                vObject = pRow["A20_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ��� ���� ȯ�޼���
                vXLine = 21;
                vXLColumn = 25;
                vObject = pRow["A20_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����ҵ� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 21;
                vXLColumn = 27;
                vObject = pRow["A20_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //------------------------------------------------------------------------------------------
                //����ҵ� �ſ�¡�� �ο���
                vXLine = 22;
                vXLColumn = 10;
                vObject = pRow["A25_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �ſ�¡�� �����޾�
                vXLine = 22;
                vXLColumn = 12;
                vObject = pRow["A25_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �ſ�¡�� �ҵ漼��
                vXLine = 22;
                vXLColumn = 16;
                vObject = pRow["A25_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �ſ�¡�� ���꼼
                vXLine = 22;
                vXLColumn = 23;
                vObject = pRow["A25_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //����ҵ� �������� �ο���
                vXLine = 23;
                vXLColumn = 10;
                vObject = pRow["A26_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �������� �����޾�
                vXLine = 23;
                vXLColumn = 12;
                vObject = pRow["A26_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �������� �ҵ漼��
                vXLine = 23;
                vXLColumn = 16;
                vObject = pRow["A26_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �������� ��Ư��
                vXLine = 23;
                vXLColumn = 20;
                vObject = pRow["A26_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� �������� ���꼼
                vXLine = 23;
                vXLColumn = 23;
                vObject = pRow["A26_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //����ҵ� ������ �ο���
                vXLine = 24;
                vXLColumn = 10;
                vObject = pRow["A30_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ �����޾�
                vXLine = 24;
                vXLColumn = 12;
                vObject = pRow["A30_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ �ҵ漼��
                vXLine = 24;
                vXLColumn = 16;
                vObject = pRow["A30_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ��Ư��
                vXLine = 24;
                vXLColumn = 20;
                vObject = pRow["A30_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ���꼼
                vXLine = 24;
                vXLColumn = 23;
                vObject = pRow["A30_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ���� ȯ�޼���
                vXLine = 24;
                vXLColumn = 25;
                vObject = pRow["A30_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 24;
                vXLColumn = 27;
                vObject = pRow["A30_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //����ҵ� ������ ���μ��� ��Ư��
                vXLine = 24;
                vXLColumn = 30;
                vObject = pRow["A30_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //----------------------------------------2017 �����߰�

                //��Ÿ�ҵ� ���ݰ��� �ο���
                vXLine = 25;
                vXLColumn = 10;
                vObject = pRow["A41_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ���ݰ��� �����޾�
                vXLine = 25;
                vXLColumn = 12;
                vObject = pRow["A41_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ���ݰ��� �ҵ漼��
                vXLine = 25;
                vXLColumn = 16;
                vObject = pRow["A41_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ���ݰ��� ���꼼
                vXLine = 25;
                vXLColumn = 23;
                vObject = pRow["A41_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ���ݰ��� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 25;
                vXLColumn = 27;
                vObject = pRow["A41_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                
                ////��Ÿ�ҵ� �����μҵ� �ſ�¡�� �ο���
                vXLine = 26;
                vXLColumn = 10;
                vObject = pRow["A43_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ� �����μҵ� �ſ�¡�� �����޾�
                vXLine = 26;
                vXLColumn = 12;
                vObject = pRow["A43_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ� �����μҵ� �ſ�¡�� �ҵ漼��
                vXLine = 26;
                vXLColumn = 16;
                vObject = pRow["A43_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ� �����μҵ� �ſ�¡�� ���꼼
                vXLine = 26;
                vXLColumn = 23;
                vObject = pRow["A43_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ� �����μҵ� �ſ�¡�� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 26;
                vXLColumn = 27;
                vObject = pRow["A43_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ�  �����μҵ� �������� �ο���
                vXLine = 27;
                vXLColumn = 10;
                vObject = pRow["A44_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ�  �����μҵ� �������� �����޾�
                vXLine = 27;
                vXLColumn = 12;
                vObject = pRow["A44_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ�  �����μҵ� �������� �ҵ漼��
                vXLine = 27;
                vXLColumn = 16;
                vObject = pRow["A44_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ�  �����μҵ� �������� ���꼼
                vXLine = 27;
                vXLColumn = 23;
                vObject = pRow["A44_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                ////��Ÿ�ҵ� �����μҵ� �������� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 27;
                vXLColumn = 27;
                vObject = pRow["A44_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));



                //----------------------------------------2017 �����߰�


                //��Ÿ�ҵ� �׿� �ο���
                vXLine = 26+2;
                vXLColumn = 10;
                vObject = pRow["A42_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� �����޾�
                vXLine = 26 + 2;
                vXLColumn = 12;
                vObject = pRow["A42_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� �ҵ漼��
                vXLine = 26 + 2;
                vXLColumn = 16;
                vObject = pRow["A42_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� ���꼼
                vXLine = 26 + 2;
                vXLColumn = 23;
                vObject = pRow["A42_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� �׿� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 26 + 2;
                vXLColumn = 27;
                vObject = pRow["A42_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ �ο���
                vXLine = 27 + 2;
                vXLColumn = 10;
                vObject = pRow["A40_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ �����޾�
                vXLine = 27 + 2;
                vXLColumn = 12;
                vObject = pRow["A40_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ �ҵ漼��
                vXLine = 27 + 2;
                vXLColumn = 16;
                vObject = pRow["A40_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ ���꼼
                vXLine = 27 + 2;
                vXLColumn = 23;
                vObject = pRow["A40_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ ���� ȯ�޼���
                vXLine = 27 + 2;
                vXLColumn = 25;
                vObject = pRow["A40_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //��Ÿ�ҵ� ������ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 27 + 2;
                vXLColumn = 27;
                vObject = pRow["A40_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //------------------------------------------------------------------------------------------
                //���ݼҵ� ���ݰ��� �ο���
                vXLine = 28 + 2;
                vXLColumn = 10;
                vObject = pRow["A48_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ���ݰ��� �����޾�
                vXLine = 28 + 2;
                vXLColumn = 12;
                vObject = pRow["A48_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ���ݰ��� �ҵ漼��
                vXLine = 28 + 2;
                vXLColumn = 16;
                vObject = pRow["A48_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ���ݰ��� ���꼼
                vXLine = 28 + 2;
                vXLColumn = 23;
                vObject = pRow["A48_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ��������(�ſ�) �ο���
                vXLine = 29 + 2;
                vXLColumn = 10;
                vObject = pRow["A45_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ��������(�ſ�) �����޾�
                vXLine = 29 + 2;
                vXLColumn = 12;
                vObject = pRow["A45_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ��������(�ſ�) �ҵ漼��
                vXLine = 29 + 2;
                vXLColumn = 16;
                vObject = pRow["A45_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ��������(�ſ�) ���꼼
                vXLine = 29 + 2;
                vXLColumn = 23;
                vObject = pRow["A45_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));


                //���ݼҵ� �������� �ο���
                vXLine = 30 + 2;
                vXLColumn = 10;
                vObject = pRow["A46_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� �������� �����޾�
                vXLine = 30 + 2;
                vXLColumn = 12;
                vObject = pRow["A46_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� �������� �ҵ漼��
                vXLine = 30 + 2;
                vXLColumn = 16;
                vObject = pRow["A46_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� �������� ���꼼
                vXLine = 30 + 2;
                vXLColumn = 23;
                vObject = pRow["A46_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ �ο���
                vXLine = 31 + 2;
                vXLColumn = 10;
                vObject = pRow["A47_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ �����޾�
                vXLine = 31 + 2;
                vXLColumn = 12;
                vObject = pRow["A47_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ �ҵ漼��
                vXLine = 31 + 2;
                vXLColumn = 16;
                vObject = pRow["A47_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ ���꼼
                vXLine = 31 + 2;
                vXLColumn = 23;
                vObject = pRow["A47_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ ���� ȯ�޼���
                vXLine = 31 + 2;
                vXLColumn = 25;
                vObject = pRow["A47_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ݼҵ� ������ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 31 + 2;
                vXLColumn = 27;
                vObject = pRow["A47_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //----------------------------------------------------------------------------
                //���ڼҵ� �ο���
                vXLine = 32 + 2;
                vXLColumn = 10;
                vObject = pRow["A50_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� �����޾�
                vXLine = 32 + 2;
                vXLColumn = 12;
                vObject = pRow["A50_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� �ҵ漼��
                vXLine = 32 + 2;
                vXLColumn = 16;
                vObject = pRow["A50_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ��Ư��
                vXLine = 32 + 2;
                vXLColumn = 20;
                vObject = pRow["A50_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ���꼼
                vXLine = 32 + 2;
                vXLColumn = 23;
                vObject = pRow["A50_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ���� ȯ�޼���
                vXLine = 32 + 2;
                vXLColumn = 25;
                vObject = pRow["A50_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 32 + 2;
                vXLColumn = 27;
                vObject = pRow["A50_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ڼҵ� ���μ��� ��Ư��
                vXLine = 32 + 2;
                vXLColumn = 30;
                vObject = pRow["A50_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //-------------------------------------------------------------------------------
                //���ҵ� �ο���
                vXLine = 33 + 2;
                vXLColumn = 10;
                vObject = pRow["A60_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� �����޾�
                vXLine = 33 + 2;
                vXLColumn = 12;
                vObject = pRow["A60_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� �ҵ漼��
                vXLine = 33 + 2;
                vXLColumn = 16;
                vObject = pRow["A60_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ��Ư��
                vXLine = 33 + 2;
                vXLColumn = 20;
                vObject = pRow["A60_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ���꼼
                vXLine = 33 + 2;
                vXLColumn = 23;
                vObject = pRow["A60_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ���� ȯ�޼���
                vXLine = 33 + 2;
                vXLColumn = 25;
                vObject = pRow["A60_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 33 + 2;
                vXLColumn = 27;
                vObject = pRow["A60_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ҵ� ���μ��� ��Ư��
                vXLine = 33 + 2;
                vXLColumn = 30;
                vObject = pRow["A60_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //----------------------------------------------------------------------------------
                //�������� �ο���
                vXLine = 34 + 2;
                vXLColumn = 10;
                vObject = pRow["A69_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� �ҵ漼��
                vXLine = 34 + 2;
                vXLColumn = 16;
                vObject = pRow["A69_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���꼼
                vXLine = 34 + 2;
                vXLColumn = 23;
                vObject = pRow["A69_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���� ȯ�޼���
                vXLine = 34 + 2;
                vXLColumn = 25;
                vObject = pRow["A69_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�������� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 34 + 2;
                vXLColumn = 27;
                vObject = pRow["A69_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //-------------------------------------------------------------------------------
                //������� �ο���
                vXLine = 35 + 2;
                vXLColumn = 10;
                vObject = pRow["A70_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� �����޾�
                vXLine = 35 + 2;
                vXLColumn = 12;
                vObject = pRow["A70_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� �ҵ漼��
                vXLine = 35 + 2;
                vXLColumn = 16;
                vObject = pRow["A70_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� ���꼼
                vXLine = 35 + 2;
                vXLColumn = 23;
                vObject = pRow["A70_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� ���� ȯ�޼���
                vXLine = 35 + 2;
                vXLColumn = 25;
                vObject = pRow["A70_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //������� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 35 + 2;
                vXLColumn = 27;
                vObject = pRow["A70_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //-------------------------------------------------------------------------------
                //���ܱ��ι��ο�õ �ο���
                vXLine = 36 + 2;
                vXLColumn = 10;
                vObject = pRow["A80_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ �����޾�
                vXLine = 36 + 2;
                vXLColumn = 12;
                vObject = pRow["A80_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ �ҵ漼��
                vXLine = 36 + 2;
                vXLColumn = 16;
                vObject = pRow["A80_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ ���꼼
                vXLine = 36 + 2;
                vXLColumn = 23;
                vObject = pRow["A80_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ ���� ȯ�޼���
                vXLine = 36 + 2;
                vXLColumn = 25;
                vObject = pRow["A80_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���ܱ��ι��ο�õ ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 36 + 2;
                vXLColumn = 27;
                vObject = pRow["A80_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //---------------------------------------------------------------------
                //�����Ű� �ҵ漼��
                vXLine = 37 + 2;
                vXLColumn = 16;
                vObject = pRow["A90_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ��Ư��
                vXLine = 37 + 2;
                vXLColumn = 20;
                vObject = pRow["A90_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ���꼼
                vXLine = 37 + 2;
                vXLColumn = 23;
                vObject = pRow["A90_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ���� ȯ�޼���
                vXLine = 37 + 2;
                vXLColumn = 25;
                vObject = pRow["A90_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 37 + 2;
                vXLColumn = 27;
                vObject = pRow["A90_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //�����Ű� ���μ��� ��Ư��
                vXLine = 37 + 2;
                vXLColumn = 30;
                vObject = pRow["A90_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //---------------------------------------------------------------------------------
                //���հ� �ο���
                vXLine = 38 + 2;
                vXLColumn = 10;
                vObject = pRow["A99_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� �����޾�
                vXLine = 38 + 2;
                vXLColumn = 12;
                vObject = pRow["A99_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� �ҵ漼��
                vXLine = 38 + 2;
                vXLColumn = 16;
                vObject = pRow["A99_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ��Ư��
                vXLine = 38 + 2;
                vXLColumn = 20;
                vObject = pRow["A99_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ���꼼
                vXLine = 38 + 2;
                vXLColumn = 23;
                vObject = pRow["A99_ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ���� ȯ�޼���
                vXLine = 38 + 2;
                vXLColumn = 25;
                vObject = pRow["A99_THIS_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ���μ��� �ҵ漼�� ���꼼 ����
                vXLine = 38 + 2;
                vXLColumn = 27;
                vObject = pRow["A99_PAY_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���հ� ���μ��� ��Ư��
                vXLine = 38 + 2;
                vXLColumn = 30;
                vObject = pRow["A99_PAY_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //--------------------------------------------------------------------------------
                //12.������ȯ�޼���
                vXLine = 43 + 2;
                vXLColumn = 2;
                vObject = pRow["RECEIVE_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //13.��ȯ�޽�û�Ѽ���
                vXLine = 43 + 2;
                vXLColumn = 6;
                vObject = pRow["ALREADY_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //14.�����ܾ�
                vXLine = 43 + 2;
                vXLColumn = 9;
                vObject = pRow["REFUND_BALANCE_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //15.�Ϲ�ȯ��
                vXLine = 43 + 2;
                vXLColumn = 12;
                vObject = pRow["GENERAL_REFUND_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //16.��Ź���(����ȸ���)
                vXLine = 43 + 2;
                vXLColumn = 15;
                vObject = pRow["FINANCIAL_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //17-1.�׹��� ȯ�޼���-����ȸ���
                vXLine = 43 + 2;
                vXLColumn = 18;
                vObject = pRow["ETC_REFUND_FINANCIAL_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //17-2.�׹��� ȯ�޼���-�պ���
                vXLine = 43 + 2;
                vXLColumn = 20;
                vObject = pRow["ETC_REFUND_MERGER_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //18.�������ȯ�޼���
                vXLine = 43 + 2;
                vXLColumn = 22;
                vObject = pRow["ADJUST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //19. ��� ������ȯ�޼���
                vXLine = 43 + 2;
                vXLColumn = 25;
                vObject = pRow["THIS_ADJUST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //20. ���� �̿� ȯ�޼���
                vXLine = 43 + 2;
                vXLColumn = 27;
                vObject = pRow["NEXT_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //21.ȯ�޽�û��
                vXLine = 43 + 2;
                vXLColumn = 30;
                vObject = pRow["REQUEST_REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //--------------------------------------------------------------
                //��������
                vXLine = 47 + 2;
                vXLColumn = 10;
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�Ű���
                vXLine = 48 + 2;
                vXLColumn = 7;
                vObject = pRow["WITHHOLDING_AGENT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������
                vXLine = 54 + 2;
                vXLColumn = 3;
                vObject = pRow["TAX_OFFICE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }


        private int LineWrite_11_SUB_01(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //�ڵ� 
                vXLColumn = 12;
                if (iString.ISNull(pRow["INCOME_SUB_CODE"]) != string.Empty)
                {
                    vObject = string.Format("{0}", pRow["INCOME_SUB_CODE"]);
                }
                else
                {
                    vObject = "";
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //�ο�
                vXLColumn = 14; 
                vObject = pRow["PERSON_CNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����޾�
                vXLColumn = 16;
                vObject = pRow["PAYMENT_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ҵ漼��
                vXLColumn = 19;
                vObject = pRow["INCOME_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư����
                vXLColumn = 22;
                vObject = pRow["SP_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //���꼼 
                vXLColumn = 24;
                vObject = pRow["ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ȯ�޼���. 
                vXLColumn = 26;
                vObject = pRow["REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���� �ҵ漼  
                vXLColumn = 28;
                vObject = pRow["FIX_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���γ�Ư�� 
                vXLColumn = 30;
                vObject = pRow["FIX_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));  

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        private int LineWrite_11_SUB_02(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //�ڵ� 
                vXLColumn = 12;
                if (iString.ISNull(pRow["INCOME_SUB_CODE"]) != string.Empty)
                {
                    vObject = string.Format("{0}", pRow["INCOME_SUB_CODE"]);
                }
                else
                {
                    vObject = "";
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //�ο�
                vXLColumn = 14;
                vObject = pRow["PERSON_CNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����޾�
                vXLColumn = 16;
                vObject = pRow["PAYMENT_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ҵ漼��
                vXLColumn = 19;
                vObject = pRow["INCOME_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư����
                vXLColumn = 22;
                vObject = pRow["SP_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���꼼 
                vXLColumn = 24;
                vObject = pRow["ADD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ȯ�޼���. 
                vXLColumn = 26;
                vObject = pRow["REFUND_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���� �ҵ漼  
                vXLColumn = 28;
                vObject = pRow["FIX_INCOME_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //���γ�Ư�� 
                vXLColumn = 30;
                vObject = pRow["FIX_SP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString.Replace("-", "��"));

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        #endregion;

        #region ----- Excel Write [CURRENCY] Method -----

        private int LineWrite2(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;

            try
            {
               //�з���ȣ                
                vObject = pRow["CLASSIFY_TYPE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 1;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                                
                //���ڵ�
                vObject = pRow["CITY_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 4;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���γ��
                vObject = pRow["SUBMIT_YYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���α���.
                vObject = pRow["SUBMIT_TYPE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����
                vObject = pRow["TAX_TYPE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����¡������.
                vXLColumn = 22;
                vObject = pRow["TAX_OFFICE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 19;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���¹�ȣ.
                vObject = pRow["TAX_ACCOUNT_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 23;
                vXLine = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 45;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ȣ.
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 4;
                vXLine = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 46;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ڹ�ȣ.
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                vXLine = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 46;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //ȸ�迬��.
                vObject = pRow["FISCAL_YEAR"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 25;
                vXLine = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 46;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������ּ�.
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 4;
                vXLine = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 48;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ȭ.
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 18;
                vXLine = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 48;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ͼӿ���.
                vObject = pRow["STD_YEAR"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                vXLine = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ͼӿ�.
                vObject = pRow["STD_MONTH"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���α���.
                vObject = pRow["PAYMENT_DUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 26;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM13"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-õ.
                vObject = pRow["INCOME_NUM12"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 6;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM11"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM9"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-õ.
                vObject = pRow["INCOME_NUM8"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM7"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM6"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 12;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM5"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-õ.
                vObject = pRow["INCOME_NUM4"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM3"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 16;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ��-��.
                vObject = pRow["INCOME_NUM1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 17;
                vXLine = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 35;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM13"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-õ.
                vObject = pRow["SP_NUM12"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 6;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM11"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM9"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-õ.
                vObject = pRow["SP_NUM8"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM7"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM6"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 12;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM5"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-õ.
                vObject = pRow["SP_NUM4"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM3"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 16;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����Ư������-��.
                vObject = pRow["SP_NUM1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 17;
                vXLine = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM13"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 5;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-õ.
                vObject = pRow["SUM_NUM12"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 6;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM11"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 7;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 8;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM9"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 9;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-õ.
                vObject = pRow["SUM_NUM8"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 10;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM7"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 11;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM6"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 12;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM5"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 13;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-õ.
                vObject = pRow["SUM_NUM4"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 14;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM3"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 15;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 16;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��-��.
                vObject = pRow["SUM_NUM1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumn = 17;
                vXLine = 18;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 38;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLine = 58;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        #endregion;

        #region ----- TOTAL AMOUNT Write Method -----

        //private int XLTOTAL_Line(int pXLine)
        //{// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��. pGDColumn : �׸��� ��ġ, pXLColumn : ���� ��ġ.
        //    int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
        //    int vXLColumnIndex = 0;

        //    string vConvertString = string.Empty;
        //    decimal vConvertDecimal = 0m;
        //    bool IsConvert = false;

        //    try
        //    { // ������ �����ؼ� Ÿ�� �� ������ ����.(
        //        mPrinting.XLActiveSheet(mTargetSheet);

        //        //�����հ�
        //        vXLColumnIndex = 14;
        //        IsConvert = IsConvertNumber(mTOT_DR_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //        }
        //        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

        //        //�뺯�հ�
        //        vXLColumnIndex = 20;
        //        IsConvert = IsConvertNumber(mTOT_CR_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,##0}", vConvertDecimal);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //        }
        //        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

        //        //-------------------------------------------------------------------
        //        vXLine = vXLine + 1;
        //        //-------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        mMessageError = ex.Message;
        //        mAppInterface.OnAppMessageEvent(mMessageError);
        //        System.Windows.Forms.Application.DoEvents();
        //    }
        //    return vXLine;
        //}

        #endregion;

        #region ----- PageNumber Write Method -----

        private void XLPageNumber(string pActiveSheet, object pPageNumber)
        {// ���������� ������Ʈ �����ϱ� ���� ������Ʈ�� ����ϰ� ��Ʈ�� �����Ѵ�.

            int vXLRow = 31; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLCol = 40;

            try
            { // ������ �����ؼ� Ÿ�� �� ������ ����.(
                mPrinting.XLActiveSheet(pActiveSheet);
                mPrinting.XLSetCell(vXLRow, vXLCol, pPageNumber);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #endregion;

        #region ----- Excel Wirte MAIN Methods ----

        public int ExcelWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pWITHHOLDING_DOC)
        {// ���� ȣ��Ǵ� �κ�.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;

            // �μ� - ��ȭ �μ� ����.
            mCopy_EndCol = 31;
            mCopy_EndRow = 53;
            mPrintingLastRow = 52;  //���� ������ �μ� ���� ����.

            try
            {
                // �����μ�Ǵ� ���.
                vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    mCopyLineSUM = CopyAndPaste(mPrinting, 1);
                    vPageRowCount = mCurrentRow - 1;    //ù�忡 ���ؼ��� ����row���� üũ.

                    vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;  //���� ����.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pWITHHOLDING_DOC.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();
                        
                        mCurrentRow = LineWrite(vRow, mCurrentRow); // ���� ��ġ �μ� �� ���� �μ��� ����.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // ������ ������ �̸� ó���� ���� ���
                            // ��������� �Ǵ� �հ踦 ǥ���Ѵ� �� ���.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //�հ�.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // ���ο� ������ üũ �� ����.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - (mPrintingLastRow + mDefaultPageRow));  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
                }

                #endregion;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
            return mPageNumber;
        }


        //2016��02�� �����//
        public int ExcelWrite_11(InfoSummit.Win.ControlAdv.ISDataAdapter pWITHHOLDING_DOC)
        {// ���� ȣ��Ǵ� �κ�.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;

            // �μ� - ��ȭ �μ� ����.
            mCopy_EndCol = 31;
            mCopy_EndRow = 57;
            mPrintingLastRow = 56;  //���� ������ �μ� ���� ����.

            try
            {
                // �����μ�Ǵ� ���.
                vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    mCopyLineSUM = CopyAndPaste(mPrinting, 1);
                    vPageRowCount = mCurrentRow - 1;    //ù�忡 ���ؼ��� ����row���� üũ.

                    vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;  //���� ����.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pWITHHOLDING_DOC.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_11(vRow, mCurrentRow); // ���� ��ġ �μ� �� ���� �μ��� ����.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // ������ ������ �̸� ó���� ���� ���
                            // ��������� �Ǵ� �հ踦 ǥ���Ѵ� �� ���.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //�հ�.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // ���ο� ������ üũ �� ����.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - (mPrintingLastRow + mDefaultPageRow));  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
                }

                #endregion;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
            return mPageNumber;
        }


        //2016��02�� �����//
        public int ExcelWrite_11_SUB(InfoSummit.Win.ControlAdv.ISDataAdapter pINCOME_SUB_01, InfoSummit.Win.ControlAdv.ISDataAdapter pINCOME_SUB_02)
        {// ���� ȣ��Ǵ� �κ�.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;

            // �μ� - ��ȭ �μ� ����.
            mCopy_EndCol = 31;
            mCopy_EndRow = 86;
            mPrintingLastRow = 56;  //���� ������ �μ� ���� ����.

            try
            {
                // �����μ�Ǵ� ���.
                vTotalRow = pINCOME_SUB_01.CurrentRows.Count;
                vTotalRow = vTotalRow + pINCOME_SUB_02.CurrentRows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    mCopyLineSUM = CopyAndPaste_SUB(mPrinting, 1);
                    vPageRowCount = mCurrentRow - 1;    //ù�忡 ���ؼ��� ����row���� üũ.

                    mCurrentRow = 65;
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pINCOME_SUB_01.CurrentRows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_11_SUB_01(vRow, mCurrentRow); // ���� ��ġ �μ� �� ���� �μ��� ����.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // ������ ������ �̸� ó���� ���� ���
                            // ��������� �Ǵ� �հ踦 ǥ���Ѵ� �� ���.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //�հ�.
                        }
                        else
                        {
                             
                        }
                    }

                    mCurrentRow = 112;
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pINCOME_SUB_02.CurrentRows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_11_SUB_02(vRow, mCurrentRow); // ���� ��ġ �μ� �� ���� �μ��� ����.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // ������ ������ �̸� ó���� ���� ���
                            // ��������� �Ǵ� �հ踦 ǥ���Ѵ� �� ���.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //�հ�.
                        }
                        else
                        {

                        }
                    }
                }

                #endregion;
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

        #region ----- �ҵ漼���μ� Excel Wirte MAIN Methods ----

        public int ExcelWrite2(InfoSummit.Win.ControlAdv.ISDataAdapter pWITHHOLDING_DOC)
        {// ���� ȣ��Ǵ� �κ�.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;
            try
            {
                // �����μ�Ǵ� ���.
                vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    mCopyLineSUM = CopyAndPaste2(mPrinting, 1);
                    vPageRowCount = mCurrentRow2 - 1;    //ù�忡 ���ؼ��� ����row���� üũ.

                    vTotalRow = pWITHHOLDING_DOC.OraSelectData.Rows.Count;  //���� ����.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pWITHHOLDING_DOC.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite2(vRow, mCurrentRow); // ���� ��ġ �μ� �� ���� �μ��� ����.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // ������ ������ �̸� ó���� ���� ���
                            // ��������� �Ǵ� �հ踦 ǥ���Ѵ� �� ���.
                            //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //�հ�.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // ���ο� ������ üũ �� ����.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - (mPrintingLastRow + mDefaultPageRow));  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
                }

                #endregion;
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

        #region ----- Foreign Currency - Excel Wirte MAIN Methods ----

        //public int ExcelWrite2(object pBalance_Date, InfoSummit.Win.ControlAdv.ISDataAdapter pPayment)
        //{// ���� ȣ��Ǵ� �κ�.

        //    string vMessage = string.Empty;

        //    int vTotalRow = 0;
        //    int vPageRowCount = 0;
        //    int vLIneRow = 0;
        //    bool vPrint_Flag = false;
        //    try
        //    {
        //        // �����μ�Ǵ� ���.
        //        vTotalRow = pPayment.OraSelectData.Rows.Count;

        //        //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
        //        //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
        //        // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.               

        //        #region ----- Line Write ----

        //        if (vTotalRow > 0)
        //        {
        //            HeaderWrite1(pBalance_Date);
        //            // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
        //            mCopyLineSUM = CopyAndPaste2(mPrinting, 1);
        //            vPageRowCount = mCurrentRow - 1;    //ù�忡 ���ؼ��� ����row���� üũ.

        //            mCurrentRow = 6;
        //            mPrinting.XLActiveSheet(mTargetSheet);
        //            //SetArray1(pGrid, out vGDColumn, out vXLColumn);
        //            foreach (System.Data.DataRow vRow in pPayment.OraSelectData.Rows)
        //            {
        //                vLIneRow++;
        //                vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow);
        //                mAppInterface.OnAppMessageEvent(vMessage);
        //                System.Windows.Forms.Application.DoEvents();

        //                //�����ڵ� ���� ���� üũ.
        //                vPrint_Flag = true;
        //                if (mAccount_Code == null || mAccount_Code == string.Empty || mIsNewPage == true)
        //                {
        //                    mMerger_Start = mCurrentRow;
        //                    mMerger_End = mCurrentRow;
        //                }
        //                else if (mAccount_Code != iString.ISNull(vRow["ACCOUNT_CODE"]))
        //                {

        //                    mPrinting.XLCellMerge(mMerger_Start, 1, mMerger_End, 4, true);
        //                    mMerger_Start = mCurrentRow;
        //                    mMerger_End = mCurrentRow;
        //                }
        //                else
        //                {
        //                    vPrint_Flag = false;
        //                    mMerger_End = mCurrentRow;
        //                }
        //                mAccount_Code = iString.ISNull(vRow["ACCOUNT_CODE"]);

        //                mCurrentRow = LineWrite2(vRow, mCurrentRow, vPrint_Flag); // ���� ��ġ �μ� �� ���� �μ��� ����.
        //                vPageRowCount = vPageRowCount + 1;

        //                if (vLIneRow == vTotalRow)
        //                {
        //                    // ������ ������ �̸� ó���� ���� ���
        //                    // ��������� �Ǵ� �հ踦 ǥ���Ѵ� �� ���.
        //                    //mCurrentRow = XLTOTAL_Line(mPageNumber * mCopy_EndRow - 4);      //�հ�.
        //                }
        //                else
        //                {
        //                    IsNewPage(vPageRowCount);   // ���ο� ������ üũ �� ����.
        //                    if (mIsNewPage == true)
        //                    {
        //                        mCurrentRow = mCurrentRow + (mCopy_EndRow - (mPrintingLastRow + mDefaultPageRow));  // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
        //                        vPageRowCount = mDefaultPageRow;
        //                    }
        //                }
        //            }
        //        }

        //        #endregion;
        //    }
        //    catch (System.Exception ex)
        //    {
        //        mMessageError = ex.Message;
        //        mPrinting.XLOpenFileClose();
        //        mPrinting.XLClose();
        //    }
        //    return mPageNumber;
        //}

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount)
        {
            int iDefaultEndRow = 1;
            if (pPageRowCount == mPrintingLastRow)
            { // pPrintingLine : ���� ��µ� ��.
                mIsNewPage = true;
                iDefaultEndRow = mCopy_EndRow - (mPrintingLastRow + 2);
                mCopyLineSUM = CopyAndPaste(mPrinting, mCurrentRow + iDefaultEndRow);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //������ ActiveSheet�� ������ ����  ������ ����
        private int CopyAndPaste(XL.XLPrint pPrinting, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;
            string vActiveSheet = mSourceSheet1;

            mPageNumber = mPageNumber + 1;
            //if (mPageNumber > 1)
            //{
            //    2��° �μ��������� �ٸ� ����� ��� ���.
            //    vActiveSheet = mSourceSheet2;   
            //}

            // page�� ǥ��.
            //XLPageNumber(pActiveSheet, mPageNumber);

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(vActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.

            return vPasteEndRow;


            //int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            //pPrinting.XLActiveSheet("SourceTab1");
            //object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            //pPrinting.XLActiveSheet("Destination");
            //object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            //pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.


            //mPageNumber++; //������ ��ȣ
            //// ������ ��ȣ ǥ��.
            ////string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            ////int vRowSTART = vCopyPrintingRowEnd - 2;
            ////int vRowEND = vCopyPrintingRowEnd - 2;
            ////int vColumnSTART = 30;
            ////int vColumnEND = 33;
            ////mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
            ////mPrinting.XLSetCell(vRowSTART, vColumnSTART, vPageNumberText); //������ ��ȣ, XLcell[��, ��]

            //return vCopySumPrintingLine;
        }

        //������ ActiveSheet�� ������ ����  ������ ����
        private int CopyAndPaste_SUB(XL.XLPrint pPrinting, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;
            string vActiveSheet = mSourceSheet2;

            mPageNumber = mPageNumber + 2;
            //if (mPageNumber > 1)
            //{
            //    2��° �μ��������� �ٸ� ����� ��� ���.
            //    vActiveSheet = mSourceSheet2;   
            //}

            // page�� ǥ��.
            //XLPageNumber(pActiveSheet, mPageNumber);

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(vActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(59, 1, vPasteEndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.

            return vPasteEndRow;


            //int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            //pPrinting.XLActiveSheet("SourceTab1");
            //object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            //pPrinting.XLActiveSheet("Destination");
            //object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            //pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.


            //mPageNumber++; //������ ��ȣ
            //// ������ ��ȣ ǥ��.
            ////string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            ////int vRowSTART = vCopyPrintingRowEnd - 2;
            ////int vRowEND = vCopyPrintingRowEnd - 2;
            ////int vColumnSTART = 30;
            ////int vColumnEND = 33;
            ////mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
            ////mPrinting.XLSetCell(vRowSTART, vColumnSTART, vPageNumberText); //������ ��ȣ, XLcell[��, ��]

            //return vCopySumPrintingLine;
        }

        private int CopyAndPaste2(XL.XLPrint pPrinting, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow2;
            string vActiveSheet = mSourceSheet1;

            mPageNumber = mPageNumber + 1;
            
            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(vActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow2, mCopy_StartCol2, mCopy_EndRow2, mCopy_EndCol2);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(mCopy_StartRow2, mCopy_StartCol2, mCopy_EndRow2, mCopy_EndCol2);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.

            return vPasteEndRow;
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //������ ActiveSheet�� ������ ����  ������ ����
        //private int CopyAndPaste2(XL.XLPrint pPrinting, int pPasteStartRow)
        //{
        //    int vPasteEndRow = pPasteStartRow + mCopy_EndRow2;
        //    string vActiveSheet = mSourceSheet1;

        //    mPageNumber = mPageNumber + 1;
        //    //if (mPageNumber > 1)
        //    //{
        //    //    2��° �μ��������� �ٸ� ����� ��� ���.
        //    //    vActiveSheet = mSourceSheet2;   
        //    //}

        //    // page�� ǥ��.
        //    //XLPageNumber(pActiveSheet, mPageNumber);

        //    //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
        //    //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
        //    pPrinting.XLActiveSheet(vActiveSheet);
        //    object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow2, mCopy_StartCol2, mCopy_EndRow2, mCopy_EndCol2);

        //    //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
        //    //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
        //    pPrinting.XLActiveSheet(mTargetSheet);
        //    object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol2, vPasteEndRow, mCopy_EndCol2);
        //    pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.

        //    return vPasteEndRow;


        //    //int vCopySumPrintingLine = pCopySumPrintingLine;

        //    //int vCopyPrintingRowSTART = vCopySumPrintingLine;
        //    //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
        //    //int vCopyPrintingRowEnd = vCopySumPrintingLine;

        //    //pPrinting.XLActiveSheet("SourceTab1");
        //    //object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
        //    //pPrinting.XLActiveSheet("Destination");
        //    //object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
        //    //pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.


        //    //mPageNumber++; //������ ��ȣ
        //    //// ������ ��ȣ ǥ��.
        //    ////string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
        //    ////int vRowSTART = vCopyPrintingRowEnd - 2;
        //    ////int vRowEND = vCopyPrintingRowEnd - 2;
        //    ////int vColumnSTART = 30;
        //    ////int vColumnEND = 33;
        //    ////mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
        //    ////mPrinting.XLSetCell(vRowSTART, vColumnSTART, vPageNumberText); //������ ��ȣ, XLcell[��, ��]

        //    //return vCopySumPrintingLine;
        //}

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            //mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
            mPrinting.XLPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            if (iString.ISNull(pSaveFileName) == string.Empty)
            {
                return;
            }

            //int vMaxNumber = MaxIncrement(pSavePath.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", pSavePath, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
            mPrinting.XLSave(pSaveFileName);

            //��ȣ�� �ּ� ó�� : ���� ��� ����.
            //if (pSaveFileName == string.Empty)
            //{
            //    return;
            //}
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = 1; //MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            //mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
