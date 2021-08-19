using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace HRMF0783
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
        private int mCopy_EndCol = 33;
        private int mCopy_EndRow = 27;
        private int mPrintingLastRow = 27;  //���� ������ �μ� ���� ����.

        private int mCurrentRow = 12;        //���� �μ�Ǵ� row ��ġ.
        private int mDefaultPageRow = 11;    //������ skip�� ����Ǵ� �⺻ PageCount �⺻��.

        // �μ�2 - �ҵ漼 ���μ� �μ� ����.
        private int mCopy_StartCol2 = 1;
        private int mCopy_StartRow2 = 1;
        private int mCopy_EndCol2 = 68;
        private int mCopy_EndRow2 = 38;
        private int mPrintingLastRow2 = 38;  //���� ������ �μ� ���� ����.

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
                //�ͼӿ�
                vObject = pRow["STD_MM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 4;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������
                vObject = pRow["PAY_SUPPLY_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 4;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ڼҵ� �ο���
                vObject = pRow["A01_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ڼҵ� ����ǥ��
                vObject = pRow["A01_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ڼҵ� ����ҵ漼
                vObject = pRow["A01_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ҵ� �ο���
                vObject = pRow["A02_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 7;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ҵ� ����ǥ��
                vObject = pRow["A02_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 7;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ҵ� ����ҵ漼
                vObject = pRow["A02_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 7;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ҵ� �ο���
                vObject = pRow["A03_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 8;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ҵ� ����ǥ��
                vObject = pRow["A03_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 8;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ҵ� ����ҵ漼
                vObject = pRow["A03_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 8;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ� �ο���
                vObject = pRow["A04_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ� ����ǥ��
                vObject = pRow["A04_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ� ����ҵ漼
                vObject = pRow["A04_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ݼҵ� �ο���
                vObject = pRow["A05_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 10;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ݼҵ� ����ǥ��
                vObject = pRow["A05_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 10;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ݼҵ� ����ҵ漼
                vObject = pRow["A05_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 10;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��Ÿ�ҵ� �ο���
                vObject = pRow["A06_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 11;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��Ÿ�ҵ� ����ǥ��
                vObject = pRow["A06_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 11;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��Ÿ�ҵ� ����ҵ漼
                vObject = pRow["A06_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 11;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����ҵ� �ο���
                vObject = pRow["A07_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 12;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����ҵ� ����ǥ��
                vObject = pRow["A07_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 12;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����ҵ� ����ҵ漼
                vObject = pRow["A07_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 12;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ܱ������κ��͹����ҵ� �ο���
                vObject = pRow["A08_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 13;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ܱ������κ��͹����ҵ� ����ǥ��
                vObject = pRow["A08_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 13;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ܱ������κ��͹����ҵ� ����ҵ漼
                vObject = pRow["A08_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 13;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���μ��� ��98�� �ο���
                vObject = pRow["A09_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 14;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���μ��� ��98�� ����ǥ��
                vObject = pRow["A09_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 14;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���μ��� ��98�� ����ҵ漼
                vObject = pRow["A09_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 14;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ҵ漼�� ��119�� �ο���
                vObject = pRow["A10_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ҵ漼�� ��119�� ����ǥ��
                vObject = pRow["A10_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ҵ漼�� ��119�� ����ҵ漼
                vObject = pRow["A10_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��������(������)����ҵ漼
                vObject = pRow["TOTAL_ADJUST_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 16;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�� �ο���
                vObject = pRow["A90_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�� ����ǥ��
                vObject = pRow["A90_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�� ����ҵ漼
                vObject = pRow["PAY_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��������
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����� �ּ�
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 20;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //��ȣ
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 21;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ڵ�Ϲ�ȣ
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 22;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ǥ��
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 23;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ȭ��ȣ
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 24;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�Ű��ϴ� �ñ���
                vObject = pRow["TAX_OFFICER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 13;
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

        #region ----- Excel Write [CURRENCY] Method -----

        private int LineWrite2(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;

            try
            {
                //����ڵ�                
                vObject = pRow["TAX_OFFICE_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 4;
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                                
                //ȸ���ڵ�
                vObject = pRow["TAX_ACCOUNT_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 4;
                vXLColumn = 16;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 62;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�������
                vObject = pRow["TAX_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 5;
                vXLColumn = 5;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���α���.
                vObject = pRow["DUE_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 5;
                vXLColumn = 16;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 39;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 62;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ȣ
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 6;
                vXLColumn = 11;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57; 
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֹ�(����)��Ϲ�ȣ
                vObject = pRow["LEGAL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 7;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ǥ��
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 8;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ڵ�Ϲ�ȣ
                vObject = pRow["VAT_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 9;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ּ�
                vObject = pRow["ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 10;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ȭ��ȣ
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 11;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ͼӳ��.
                vObject = pRow["STD_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 12;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 47;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�Ű��ϴ� �ñ���.
                vObject = pRow["TAX_OFFICER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 13;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ξ�
                vObject = pRow["PAY_LOCAL_TAX_KOR"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 14;
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 53;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ڼҵ� �ο���
                vObject = pRow["A01_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 16;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ڼҵ� ����ǥ��
                vObject = pRow["A01_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 16;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ڼҵ� ����ҵ漼
                vObject = pRow["A01_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 16;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ҵ� �ο���
                vObject = pRow["A02_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ҵ� ����ǥ��
                vObject = pRow["A02_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ҵ� ����ҵ漼
                vObject = pRow["A02_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ҵ� �ο���
                vObject = pRow["A03_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ҵ� ����ǥ��
                vObject = pRow["A03_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ҵ� ����ҵ漼
                vObject = pRow["A03_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //�ٷμҵ� �ο���
                vObject = pRow["A04_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 19;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ� ����ǥ��
                vObject = pRow["A04_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 19;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ٷμҵ� ����ҵ漼
                vObject = pRow["A04_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 19;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ݼҵ� �ο���
                vObject = pRow["A05_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 20;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ݼҵ� ����ǥ��
                vObject = pRow["A05_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 20;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���ݼҵ� ����ҵ漼
                vObject = pRow["A05_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 20;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��Ÿ�ҵ� �ο���
                vObject = pRow["A06_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 21;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��Ÿ�ҵ� ����ǥ��
                vObject = pRow["A06_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 21;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��Ÿ�ҵ� ����ҵ漼
                vObject = pRow["A06_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 21;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����ҵ� �ο���
                vObject = pRow["A07_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 22;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����ҵ� ����ǥ��
                vObject = pRow["A07_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 22;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����ҵ� ����ҵ漼
                vObject = pRow["A07_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 22;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ܱ������κ��� �����ҵ� �ο���
                vObject = pRow["A08_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 23;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ܱ������κ��� �����ҵ� ����ǥ��
                vObject = pRow["A08_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 23;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ܱ������κ��� �����ҵ� ����ҵ漼
                vObject = pRow["A08_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 23;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���μ��� ��98�� �ο���
                vObject = pRow["A09_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���μ��� ��98�� ����ǥ��
                vObject = pRow["A09_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���μ��� ��98�� ����ҵ漼
                vObject = pRow["A09_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ҵ漼�� ��119�� �ο���
                vObject = pRow["A10_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 27;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ҵ漼�� ��119�� ����ǥ��
                vObject = pRow["A10_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 27;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ҵ漼�� ��119�� ����ҵ漼
                vObject = pRow["A10_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 27;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��������(������)
                vObject = pRow["TOTAL_ADJUST_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 29;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�� �ο���
                vObject = pRow["A90_PERSON_CNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 30;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�� ����ǥ��
                vObject = pRow["A90_STD_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 30;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 57;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�� ����ҵ漼
                vObject = pRow["PAY_LOCAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 30;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 40;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 63;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��������.
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 32;
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                vXLine = 33;
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���� �Ǵ� ��.
                vObject = pRow["REPORT_CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 33;
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����
                vObject = pRow["TAX_OFFIECER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 37;
                vXLColumn = 5;
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
