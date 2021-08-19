using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace HRMF0786
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
        private int mCopy_EndCol = 32;
        private int mCopy_EndRow = 165;
        private int mPrintingLastRow = 165;  //���� ������ �μ� ���� ����.

        private int mCurrentRow = 12;        //���� �μ�Ǵ� row ��ġ.
        private int mDefaultPageRow = 11;    //������ skip�� ����Ǵ� �⺻ PageCount �⺻��.

        // �μ�2 - �ҵ漼 ���μ� �μ� ����.
        private int mCopy_StartCol2 = 1;
        private int mCopy_StartRow2 = 1;
        private int mCopy_EndCol2 = 60;
        private int mCopy_EndRow2 = 25;
        private int mPrintingLastRow2 = 23;  //���� ������ �μ� ���� ����.

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
                //������(��ȣ)
                vObject = pRow["CORP_SITE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 5;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֹ�(����)��Ϲ�ȣ
                vObject = pRow["LEGAL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���� ������
                vObject = pRow["CORP_ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 7;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //����� ������
                vObject = pRow["OPERATING_UNIT_ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 10;
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
                vXLine = 8;
                vXLColumn = 21;
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
                vXLine = 11;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ѽ���ȣ
                vObject = pRow["FAX_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 11;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                ////�Ű�����.
                //vObject = pRow["STD_REPORT_TITLE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vConvertString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vConvertString = null;
                //}
                //vXLine = 9;
                //vXLColumn = 1;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����� �ο�
                vObject = pRow["PERSON_COUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱ��ϳⰣ �������޿��Ѿ��� ����ձݾ�
                vObject = pRow["YEAR_AVG_SALARY_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�޿��Ѿ�
                vObject = pRow["TOTAL_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�������� �޿�
                vObject = pRow["TAX_FREE_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����޿��Ѿ�
                vObject = pRow["PAYMENT_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //14
                vObject = pRow["DED_PRE_WORKER_COUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //15
                vObject = pRow["DED_THIS_SALARY_AMT"];
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

                //16
                vObject = pRow["DED_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 22;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //17�����ǥ
                vObject = pRow["COMP_STD_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 28;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //���⼼��18
                vObject = pRow["COMP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 28;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                

                //���Ű��꼼
                vObject = pRow["BAD_REPORT_ADDITION_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 28;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                ///���ҽŰ��꼼 
                vObject = pRow["BAD_SMALL_PAY_ADDITION_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 29;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���κҼ��ǰ��꼼
                vObject = pRow["BAD_PAY_ADDITION_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 30;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //19.���꼼 ��
                vObject = pRow["TAX_ADDITION_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 31;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //20�Ű����հ�
                vObject = pRow["TOTAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 32;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�Ű���
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 38;
                vXLColumn = 20;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //¡����
                vObject = pRow["TAX_OFFICER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 39;
                vXLColumn = 3;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������
                vObject = pRow["RECEIPT_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 46;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������-����

                //�Ű�����.
                vObject = pRow["STD_REPORT_TITLE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 37;
                vXLColumn = 14;
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

        private int LineWrite6(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //������(��ȣ)
                vObject = pRow["CORP_SITE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 60;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 61;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֹ�(����)��Ϲ�ȣ
                vObject = pRow["LEGAL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 61;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                //����� ������
                vObject = pRow["OPERATING_UNIT_ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 62;
                vXLColumn = 10;
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
                vXLine = 63;
                vXLColumn = 21;
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
                vXLine = 64;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ѽ���ȣ
                vObject = pRow["FAX_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 64;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //��
                vObject = pRow["PERSON_COUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 67;
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //��ð��������

                vObject = pRow["REGULAR_WORKER_COUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 67;
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //���ð��������
                vObject = pRow["DAY_WORKER_COUNT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 67;
                vXLColumn = 19;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���
                vObject = pRow["DESCRIPTION"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 67;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //�Ű�����.
                vObject = pRow["STD_REPORT_TITLE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 69;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                //---------------------�ֱ� 12������ ���޿� �Ѿ�------------------------
                //
               //-----��
                //�ֱٱ޿� 1 
                vObject = pRow["MONTH_SALARY_1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 88;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 2
                vObject = pRow["MONTH_SALARY_2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 88;
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);



                //�ֱٱ޿� 3
                vObject = pRow["MONTH_SALARY_3"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 88;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);



                //�ֱٱ޿� 4
                vObject = pRow["MONTH_SALARY_4"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 88;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //�ֱٱ޿� 5
                vObject = pRow["MONTH_SALARY_5"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 88;
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 6
                vObject = pRow["MONTH_SALARY_6"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 88;
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //�ֱٱ޿� 7
                vObject = pRow["MONTH_SALARY_7"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 92;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 8
                vObject = pRow["MONTH_SALARY_8"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 92;
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);



                //�ֱٱ޿� 9
                vObject = pRow["MONTH_SALARY_9"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 92;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);



                //�ֱٱ޿� 10
                vObject = pRow["MONTH_SALARY_10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 92;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //�ֱٱ޿� 11
                vObject = pRow["MONTH_SALARY_11"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 92;
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 12
                vObject = pRow["MONTH_SALARY_12"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 92;
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //------------------------------------------------------------
                
                //�ֱ��ϳⰣ �������޿��Ѿ��� ����ձݾ� (�ݾ�)
                vObject = pRow["YEAR_AVG_SALARY_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 90;
                vXLColumn = 4;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 1 
                vObject = pRow["MONTH_SALARY_AMT_1"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 90;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 2
                vObject = pRow["MONTH_SALARY_AMT_2"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 90;
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);



                //�ֱٱ޿� 3
                vObject = pRow["MONTH_SALARY_AMT_3"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 90;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);



                //�ֱٱ޿� 4
                vObject = pRow["MONTH_SALARY_AMT_4"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 90;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //�ֱٱ޿� 5
                vObject = pRow["MONTH_SALARY_AMT_5"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 90;
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 6
                vObject = pRow["MONTH_SALARY_AMT_6"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 90;
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                

                //�ֱٱ޿� 7
                vObject = pRow["MONTH_SALARY_AMT_7"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 94;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 8
                vObject = pRow["MONTH_SALARY_AMT_8"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 94;
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);



                //�ֱٱ޿� 9
                vObject = pRow["MONTH_SALARY_AMT_9"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 94;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);



                //�ֱٱ޿� 10
                vObject = pRow["MONTH_SALARY_AMT_10"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 94;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //�ֱٱ޿� 11
                vObject = pRow["MONTH_SALARY_AMT_11"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 94;
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱٱ޿� 12
                vObject = pRow["MONTH_SALARY_AMT_12"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 94;
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //---------------------���� ���� ���� ������ �� -----------------------
                //��
                vObject = pRow["TOTAL_PERSON_COUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //1
                vObject = pRow["MONTH_PERSON_COUNT_1"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //2
                vObject = pRow["MONTH_PERSON_COUNT_2"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 11;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //3
                vObject = pRow["MONTH_PERSON_COUNT_3"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 13;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //4
                vObject = pRow["MONTH_PERSON_COUNT_4"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //5
                vObject = pRow["MONTH_PERSON_COUNT_5"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn =17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //6
                vObject = pRow["MONTH_PERSON_COUNT_6"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 19;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //7
                vObject = pRow["MONTH_PERSON_COUNT_7"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //8
                vObject = pRow["MONTH_PERSON_COUNT_8"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 23;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //9
                vObject = pRow["MONTH_PERSON_COUNT_9"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 25;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //10
                vObject = pRow["MONTH_PERSON_COUNT_10"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //11
                vObject = pRow["MONTH_PERSON_COUNT_11"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //12
                vObject = pRow["MONTH_PERSON_COUNT_12"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 102;
                vXLColumn = 31;
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
        private int LineWrite3(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {

                //������� ���� 
                vObject = pRow["OFFICE_TAX_DOC_ITEM_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                //vXLine = pXLine;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine , vXLColumn, vConvertString);

                //������� �޿���
                vObject = pRow["ITEM_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                //vXLine = 18;
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine , vXLColumn, vConvertString);


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
        private int LineWrite4(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
               
                //������� ���� 
                vObject = pRow["OFFICE_TAX_DOC_ITEM_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                //vXLine = pXLine;
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������� �޿���
                vObject = pRow["ITEM_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                //vXLine = 18;
                vXLColumn = 20;
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
        

        private int LineWrite5(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                //������(��ȣ)
                vObject = pRow["CORP_SITE_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 5;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����
                vObject = pRow["PRESIDENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֹ�(����)��Ϲ�ȣ
                vObject = pRow["LEGAL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 6;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���� ������
                vObject = pRow["CORP_ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 7;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //����� ������
                vObject = pRow["OPERATING_UNIT_ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 10;
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
                vXLine = 8;
                vXLColumn = 21;
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
                vXLine = 11;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ѽ���ȣ
                vObject = pRow["FAX_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 11;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�Ű�����.
                vObject = pRow["STD_REPORT_TITLE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 9;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����� �ο�
                vObject = pRow["PERSON_COUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ֱ��ϳⰣ �������޿��Ѿ��� ����ձݾ�
                vObject = pRow["YEAR_AVG_SALARY_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�޿��Ѿ�
                vObject = pRow["TOTAL_PAYMENT_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�������� �޿�
                vObject = pRow["TAX_FREE_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                                
                //�����޿��Ѿ�
                vObject = pRow["PAYMENT_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 18;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //14
                vObject = pRow["DED_PRE_WORKER_COUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //15
                vObject = pRow["DED_THIS_SALARY_AMT"];
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

                //16
                vObject = pRow["DED_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 25;
                vXLColumn = 22;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);


                //17�����ǥ
                vObject = pRow["COMP_STD_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 28;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                
                
                //���⼼��18
                vObject = pRow["COMP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 28;
                vXLColumn = 17;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���⼼��
                vObject = pRow["COMP_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 14;
                vXLColumn = 7;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���Ű��꼼
                vObject = pRow["BAD_REPORT_ADDITION_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 28;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                ///���ҽŰ��꼼 
                vObject = pRow["BAD_SMALL_PAY_ADDITION_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 29;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���κҼ��ǰ��꼼
                vObject = pRow["BAD_PAY_ADDITION_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 30;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //19.���꼼 ��
                vObject = pRow["TAX_ADDITION_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 31;
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //20�Ű����հ�
                vObject = pRow["TOTAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 32;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�Ű���
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 38;
                vXLColumn = 20;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //¡����
                vObject = pRow["TAX_OFFICER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 39;
                vXLColumn = 3;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������
                vObject = pRow["RECEIPT_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 46;
                vXLColumn = 15;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������-����
                vObject = pRow["RECEIPT_AGENT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 48;
                vXLColumn = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������-�ּ�
                vObject = pRow["RECEIPT_ADDRESS"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 48;
                vXLColumn = 19;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //������-���2
                vObject = pRow["RECEIPT_YYYYMM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 50;
                vXLColumn = 4;
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
                vXLColumn = 6;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 26; 
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 46; 
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
                vXLColumn = 15;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 35;    
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 55;    
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
                vXLColumn = 6;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 26; 
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 46; 
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
                vXLColumn = 15;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 35;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 55;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����������
                vObject = pRow["ADDRESS"];
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
                vXLColumn = 31;   
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;   
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���θ�.
                vObject = pRow["CORP_NAME"];
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
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ǥ�ڼ���
                vObject = pRow["PRESIDENT_NAME"];
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
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���� ��Ϲ�ȣ.
                vObject = pRow["LEGAL_NUMBER"];
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
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //����ڵ�Ϲ�ȣ.
                vObject = pRow["VAT_NUMBER"];
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
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��ȭ��ȣ.
                vObject = pRow["TEL_NUMBER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 12;
                vXLColumn = 11;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 31;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 51;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�ο�.
                vObject = pRow["PERSON_COUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 46;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�������ܱ޿�.
                vObject = pRow["TAX_FREE_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 49;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�����޿�
                vObject = pRow["PAYMENT_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 15;
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 54;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���μ���.
                vObject = pRow["KOR_TOTAL_TAX_AMT"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 16;
                vXLColumn = 8;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 28;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 48;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //���μ���-����.
                vObject = pRow["TOTAL_TAX_AMT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = null;
                }
                vXLine = 17;
                vXLColumn = 10;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 30;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 50;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //��������
                vObject = pRow["SUBMIT_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 20;
                vXLColumn = 9;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 29;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                vXLColumn = 49;
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�Ű���.
                vObject = pRow["CORP_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 21;
                vXLColumn = 6;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�������.
                vObject = pRow["TAX_OFFIECER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 21;
                vXLColumn = 26;                
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //�������.
                vObject = pRow["TAX_OFFIECER"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLine = 24;
                vXLColumn = 6;
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

        public int ExcelWrite( InfoSummit.Win.ControlAdv.ISDataAdapter pOFFICE_TAX_DOC
                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pOFFICE_TAX_DOC_SALARY
                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pSALARY_ITEM
                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pSALARY_TAX_FREE_ITEM
            )
        {// ���� ȣ��Ǵ� �κ�.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vTotalRow2 = 0;
            int vTotalRow3 = 0;
            int vTotalRow4 = 0;
            int vTotalRow5 = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;
            try
            {
                // �����μ�Ǵ� ���.
                vTotalRow = pOFFICE_TAX_DOC.OraSelectData.Rows.Count;
                vTotalRow2 = pOFFICE_TAX_DOC_SALARY.OraSelectData.Rows.Count;
                vTotalRow3 = pSALARY_ITEM.OraSelectData.Rows.Count;
                vTotalRow4 = pSALARY_TAX_FREE_ITEM.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    mCopyLineSUM = CopyAndPaste(mPrinting, 1);
                    vPageRowCount = mCurrentRow - 1;    //ù�忡 ���ؼ��� ����row���� üũ.

                    vTotalRow = pOFFICE_TAX_DOC.OraSelectData.Rows.Count;  //���� ����.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pOFFICE_TAX_DOC.OraSelectData.Rows)
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
                    //�޿��Ѱ�ǥ ���
                    foreach (System.Data.DataRow vRow in pOFFICE_TAX_DOC_SALARY.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow2);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite6(vRow, mCurrentRow); // ���� ��ġ �μ� �� ���� �μ��� ����.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow2)
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
                    mCurrentRow = 75;
                    //���� 
                    foreach (System.Data.DataRow vRow in pSALARY_ITEM.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow3);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite3(vRow, mCurrentRow); // ���� ��ġ �μ� �� ���� �μ��� ����.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow3)
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
                    mCurrentRow = 75;
                    //�����
                    foreach (System.Data.DataRow vRow in pSALARY_TAX_FREE_ITEM.OraSelectData.Rows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vLIneRow, vTotalRow4);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite4(vRow, mCurrentRow); // ���� ��ġ �μ� �� ���� �μ��� ����.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow4)
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

        public int ExcelWrite2(InfoSummit.Win.ControlAdv.ISDataAdapter pOFFICE_TAX)
        {// ���� ȣ��Ǵ� �κ�.

            string vMessage = string.Empty;

            int vTotalRow = 0;
            int vPageRowCount = 0;
            int vLIneRow = 0;
            try
            {
                // �����μ�Ǵ� ���.
                vTotalRow = pOFFICE_TAX.OraSelectData.Rows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    mCopyLineSUM = CopyAndPaste2(mPrinting, 1);
                    vPageRowCount = mCurrentRow2 - 1;    //ù�忡 ���ؼ��� ����row���� üũ.

                    vTotalRow = pOFFICE_TAX.OraSelectData.Rows.Count;  //���� ����.
                    mPrinting.XLActiveSheet(mTargetSheet);
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pOFFICE_TAX.OraSelectData.Rows)
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
