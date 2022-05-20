using System;

namespace HRMF0562
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        // ��Ʈ�� ����.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "Source1";
        private string mSourceSheet2 = "Source2";

        private string mMessageError = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0; 

        private string mXLOpenFileName = string.Empty;

        private bool mIsNewPage = false;    // ù ������ üũ.
        private int mMaxPrintPage = 32;     //�ѹ��� �μ��ϴ� ����ȭ ��������.
        private int mPrintPage = 0;         //�μ�ż�.

        // �μ�� ���ο� �հ�.
        private int mCopyLineSUM = 0;

        // �μ� 1���� �ִ� �μ�����.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingLastRow = 0;  //���� �μ� ����.

        private int mCurrentRow = 0;       //���� �μ�Ǵ� row ��ġ.
        private int mCurrentCol = 0;       //���� �μ�Ǵ� row ��ġ.
        private int mDefaultPageRow = 0;    // ������ ������ PageCount �⺻��.
           
        private int mPrintingLineSTART = 1;  //Line
        
        private int mIncrementCopyMAX = 70;  //����Ǿ��� ���� ����

        private int mCopyColumnSTART = 1; //����Ǿ�  �� �� ���� ��
        private int mCopyColumnEND = 45;  //������ ���õ� ��Ʈ�� ����Ǿ��� �� �� ��ġ

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
            pGDColumn = new int[2];
            pXLColumn = new int[2];

            pGDColumn[0] = pTable.Columns.IndexOf("ALLOWANCE_NAME");       //�޿� ���޸�
            pGDColumn[1] = pTable.Columns.IndexOf("ALLOWANCE_AMOUNT");     //�޿� ���ޱݾ�

            pXLColumn[0] = 6;    //�޿� ���޸�
            pXLColumn[1] = 15;   //�޿� ���޸��
        }

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn)
        {
            pGDColumn = new int[23];

            pGDColumn[0] = pGrid.GetColumnToIndex("WAGE_TYPE_NAME");            //Title
            pGDColumn[1] = pGrid.GetColumnToIndex("DEPT_NAME");                 //�μ�
            pGDColumn[2] = pGrid.GetColumnToIndex("POST_NAME");                 //����
            pGDColumn[3] = pGrid.GetColumnToIndex("PERSON_NUM");                //���
            pGDColumn[4] = pGrid.GetColumnToIndex("NAME");                      //�̸�

            pGDColumn[5] = pGrid.GetColumnToIndex("JOB_CLASS_NAME");            //����
            pGDColumn[6] = pGrid.GetColumnToIndex("SUPPLY_DATE");               //������
            pGDColumn[7] = pGrid.GetColumnToIndex("BANK_NAME");                 //�Ա�����
            pGDColumn[8] = pGrid.GetColumnToIndex("BANK_ACCOUNTS");             //�Աݰ���

            pGDColumn[9] = pGrid.GetColumnToIndex("BASIC_AMOUNT");              //�⺻��
            pGDColumn[10] = pGrid.GetColumnToIndex("BASIC_TIME_AMOUNT");            //�ñ�
            pGDColumn[11] = pGrid.GetColumnToIndex("GENERAL_HOURLY_AMOUNT");    //���ñ�

            pGDColumn[12] = pGrid.GetColumnToIndex("TOT_PAY_DED_AMOUNT");       //�޿� �Ѱ�����
            pGDColumn[13] = pGrid.GetColumnToIndex("TOT_PAY_SUP_AMOUNT");       //�޿� �����޾�
            pGDColumn[14] = pGrid.GetColumnToIndex("REAL_PAY_AMOUNT");          //�޿� �����޾�

            pGDColumn[15] = pGrid.GetColumnToIndex("TOT_BONUS_DED_AMOUNT");     //�� �Ѱ�����
            pGDColumn[16] = pGrid.GetColumnToIndex("TOT_BONUS_SUP_AMOUNT");     //�� �����޾�
            pGDColumn[17] = pGrid.GetColumnToIndex("REAL_BONUS_AMOUNT");        //�� �����޾�

            pGDColumn[18] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");        //�����޾�
            pGDColumn[19] = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");           //�Ѱ�����
            pGDColumn[20] = pGrid.GetColumnToIndex("REAL_AMOUNT");              //�� �����޾�
            pGDColumn[21] = pGrid.GetColumnToIndex("DESCRIPTION");              //���
            pGDColumn[22] = pGrid.GetColumnToIndex("CORP_NAME");                //ȸ���

        }

        #endregion;

        #region ----- Array Set 2 ----

        private void SetArray2(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[2];
            pXLColumn = new int[2];

            pGDColumn[0] = pTable.Columns.IndexOf("DEDUCTION_NAME");       //�޿� ������
            pGDColumn[1] = pTable.Columns.IndexOf("DEDUCTION_AMOUNT");     //�޿� �����ݾ�

            pXLColumn[0] = 25;   //�޿� ������
            pXLColumn[1] = 34;   //�޿� �����ݾ�
        }

        #endregion;

        #region ----- Array Set 3 ----

        private void SetArray3(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[2];
            pXLColumn = new int[2];

            pGDColumn[0] = pTable.Columns.IndexOf("ALLOWANCE_NAME");       //�� ���޸�
            pGDColumn[1] = pTable.Columns.IndexOf("ALLOWANCE_AMOUNT");     //�� ���ޱݾ�

            pXLColumn[0] = 6;    //�� ���޸�
            pXLColumn[1] = 15;   //�� ���޸��
        }

        #endregion;

        #region ----- Array Set 4 ----

        private void SetArray4(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[2];
            pXLColumn = new int[2];

            pGDColumn[0] = pTable.Columns.IndexOf("DEDUCTION_NAME");       //�� ������
            pGDColumn[1] = pTable.Columns.IndexOf("DEDUCTION_AMOUNT");     //�� �����ݾ�

            pXLColumn[0] = 25;   //�� ������
            pXLColumn[1] = 34;   //�� �����ݾ�
        }

        #endregion;

        #region ----- Array Set 5 ----

        private void SetArray5(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[16];
            pXLColumn = new int[16];

            pGDColumn[0] = pTable.Columns.IndexOf("OVER_TIME");            //����(����)
            pGDColumn[1] = pTable.Columns.IndexOf("NIGHT_BONUS_TIME");     //�߰�(����)
            pGDColumn[2] = pTable.Columns.IndexOf("LATE_TIME");            //���°���(����)
            pGDColumn[3] = pTable.Columns.IndexOf("HOLY_1_TIME");          //������-�ٹ�
            pGDColumn[4] = pTable.Columns.IndexOf("HOLY_1_OT");            //������-����
            pGDColumn[5] = pTable.Columns.IndexOf("HOLY_1_NIGHT");         //������-�߰�
            pGDColumn[6] = pTable.Columns.IndexOf("HOLY_0_TIME");          //������-�ٹ�
            pGDColumn[7] = pTable.Columns.IndexOf("HOLY_0_OT");            //������-����
            pGDColumn[8] = pTable.Columns.IndexOf("HOLY_0_NIGHT");         //������-�߰�
            pGDColumn[9] = pTable.Columns.IndexOf("TOTAL_ATT_DAY");        //�ٹ�(�ΰ�����)
            pGDColumn[10] = pTable.Columns.IndexOf("DUTY_30");             //����(�ΰ�����)
            pGDColumn[11] = pTable.Columns.IndexOf("S_HOLY_1_COUNT");      //����(�ΰ�����)
            pGDColumn[12] = pTable.Columns.IndexOf("HOLY_1_COUNT");        //����(�ΰ�����)
            pGDColumn[13] = pTable.Columns.IndexOf("HOLY_0_COUNT");        //����(�ΰ�����)
            pGDColumn[14] = pTable.Columns.IndexOf("TOT_DED_COUNT");       //�̱ٹ�(�ΰ�����)
            pGDColumn[15] = pTable.Columns.IndexOf("WEEKLY_DED_COUNT");    //������(�ΰ�����)

            pXLColumn[0] = 12;   //����(����)
            pXLColumn[1] = 16;   //�߰�(����)
            pXLColumn[2] = 20;   //���°���(����)
            pXLColumn[3] = 8;    //������-�ٹ�
            pXLColumn[4] = 12;   //������-����
            pXLColumn[5] = 16;   //������-�߰�
            pXLColumn[6] = 8;    //������-�ٹ�
            pXLColumn[7] = 12;   //������-����
            pXLColumn[8] = 16;   //������-�߰�
            pXLColumn[9] = 4;    //�ٹ�(�ΰ�����)
            pXLColumn[10] = 8;   //����(�ΰ�����)
            pXLColumn[11] = 12;  //����(�ΰ�����)
            pXLColumn[12] = 16;  //����(�ΰ�����)
            pXLColumn[13] = 20;  //����(�ΰ�����)
            pXLColumn[14] = 24;  //�̱ٹ�(�ΰ�����)
            pXLColumn[15] = 28;  //������(�ΰ�����)
        }

        private void SetArray5_BHK(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[17];
            pXLColumn = new int[17];

            pGDColumn[0] = pTable.Columns.IndexOf("TOTAL_ATT_TIME");        //�⺻�ٹ��ð� 
            pGDColumn[1] = pTable.Columns.IndexOf("OVER_TIME");             //����ٹ�  
            pGDColumn[2] = pTable.Columns.IndexOf("NIGHT_BONUS_TIME");      //�߰��ٹ�
            pGDColumn[3] = pTable.Columns.IndexOf("HOLIDAY_TIME");          //���ϱٹ�
            pGDColumn[4] = pTable.Columns.IndexOf("HOLIDAY_OVER_TIME");     //���Ͽ���ٹ� 
            pGDColumn[5] = pTable.Columns.IndexOf("LATE_TIME");             //���°��� 
            pGDColumn[6] = pTable.Columns.IndexOf("CREATION_NUM");          //�߻����� 
            pGDColumn[7] = pTable.Columns.IndexOf("REMAIN_NUM");            //�ܿ����� 
            pGDColumn[8] = pTable.Columns.IndexOf("BASE_AMOUNT");           //�⺻�� 
            pGDColumn[9] = pTable.Columns.IndexOf("TOTAL_ATT_DAY");         //�ٹ��ϼ� 
            pGDColumn[10] = pTable.Columns.IndexOf("S_HOLY_1_COUNT");       //�����ϼ� 
            pGDColumn[11] = pTable.Columns.IndexOf("H_HOLY_1_COUNT");       //�����ϼ� 
            pGDColumn[12] = pTable.Columns.IndexOf("HOLY_0_COUNT");         //�����ϼ� 
            pGDColumn[13] = pTable.Columns.IndexOf("DUTY_11");              //��� 
            pGDColumn[14] = pTable.Columns.IndexOf("LATE_DED_COUNT");       //����/����ȸ�� 
            pGDColumn[15] = pTable.Columns.IndexOf("THIS_DUTY_20_NUM");     //�����뿬����
            pGDColumn[16] = pTable.Columns.IndexOf("USE_NUM");              //������뿬����

            pXLColumn[0] = 3;       //�⺻�ٹ��ð�
            pXLColumn[1] = 10;       //����ٹ�
            pXLColumn[2] = 17;      //�߰��ٹ�
            pXLColumn[3] = 24;      //���ϱٹ�
            pXLColumn[4] = 31;      //���Ͽ���ٹ�
            pXLColumn[5] = 38;      //���°���
            pXLColumn[6] = 33;      //�߻����� 
            pXLColumn[7] = 38;      //�ܿ�����
            pXLColumn[8] = 3;       //�⺻��
            pXLColumn[9] = 19;       //�ٹ��ϼ�
            pXLColumn[10] = 24;     //�����ϼ� 
            pXLColumn[11] = 15;     //�����ϼ� 
            pXLColumn[12] = 29;     //�����ϼ� 
            pXLColumn[13] = 34;     //���
            pXLColumn[14] = 40;     //����/����ȸ�� 
            pXLColumn[15] = 33;     //�����뿬����
            pXLColumn[16] = 38;     //������뿬����
        }

        #endregion;

        #region ----- Array Set 6 ----

        private void SetArray6(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[27];
            pXLColumn = new int[26];

            pGDColumn[0] = pGrid.GetColumnToIndex("WAGE_TYPE_NAME");            //Title
            pGDColumn[1] = pGrid.GetColumnToIndex("DEPT_NAME");                 //�μ�
            pGDColumn[2] = pGrid.GetColumnToIndex("POST_NAME");                 //����
            pGDColumn[3] = pGrid.GetColumnToIndex("PERSON_NUM");                //���
            pGDColumn[4] = pGrid.GetColumnToIndex("NAME");                      //�̸�

            pGDColumn[5] = pGrid.GetColumnToIndex("NAME");                      //����
            pGDColumn[6] = pGrid.GetColumnToIndex("PERSON_NUM");                //���
            pGDColumn[7] = pGrid.GetColumnToIndex("DEPT_NAME");                 //�μ�
            pGDColumn[8] = pGrid.GetColumnToIndex("POST_NAME");                 //����
            pGDColumn[9] = pGrid.GetColumnToIndex("JOB_CLASS_NAME");            //����
            pGDColumn[10] = pGrid.GetColumnToIndex("SUPPLY_DATE");              //������
            pGDColumn[11] = pGrid.GetColumnToIndex("BANK_NAME");                //�Ա�����
            pGDColumn[12] = pGrid.GetColumnToIndex("BANK_ACCOUNTS");            //�Աݰ���

            pGDColumn[13] = pGrid.GetColumnToIndex("BASIC_AMOUNT");             //�⺻��
            pGDColumn[14] = pGrid.GetColumnToIndex("HOURLY_AMOUNT");            //�ñ�
            pGDColumn[15] = pGrid.GetColumnToIndex("GENERAL_HOURLY_AMOUNT");    //���ñ�

            pGDColumn[16] = pGrid.GetColumnToIndex("TOT_PAY_DED_AMOUNT");       //�޿� �Ѱ�����
            pGDColumn[17] = pGrid.GetColumnToIndex("TOT_PAY_SUP_AMOUNT");       //�޿� �����޾�
            pGDColumn[18] = pGrid.GetColumnToIndex("REAL_PAY_AMOUNT");          //�޿� �����޾�

            pGDColumn[19] = pGrid.GetColumnToIndex("TOT_BONUS_DED_AMOUNT");     //�� �Ѱ�����
            pGDColumn[20] = pGrid.GetColumnToIndex("TOT_BONUS_SUP_AMOUNT");     //�� �����޾�
            pGDColumn[21] = pGrid.GetColumnToIndex("REAL_BONUS_AMOUNT");        //�� �����޾�

            pGDColumn[22] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");        //�����޾�
            pGDColumn[23] = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");           //�Ѱ�����
            pGDColumn[24] = pGrid.GetColumnToIndex("REAL_AMOUNT");              //�� �����޾�

            pGDColumn[25] = pGrid.GetColumnToIndex("DESCRIPTION");              //���
            pGDColumn[26] = pGrid.GetColumnToIndex("NOTIFICATION");             //�˸�

            pXLColumn[0] = 4;       //Title
            pXLColumn[1] = 8;       //�μ�
            pXLColumn[2] = 8;       //����
            pXLColumn[3] = 8;       //���
            pXLColumn[4] = 8;       //�̸�

            pXLColumn[5] = 9;       //����
            pXLColumn[6] = 22;      //���
            pXLColumn[7] = 36;      //�μ�
            pXLColumn[8] = 9;       //����
            pXLColumn[9] = 22;      //����
            pXLColumn[10] = 36;     //������
            pXLColumn[11] = 9;      //�Ա�����
            pXLColumn[12] = 22;     //�Աݰ���

            pXLColumn[13] = 32;     //�⺻��
            pXLColumn[14] = 36;     //�ñ�
            pXLColumn[15] = 40;     //���ñ�

            pXLColumn[16] = 34;     //�޿� �Ѱ�����
            pXLColumn[17] = 15;     //�޿� �����޾�
            pXLColumn[18] = 34;     //�޿� �����޾�

            pXLColumn[19] = 34;     //�� �Ѱ�����
            pXLColumn[20] = 15;     //�� �����޾�
            pXLColumn[21] = 34;     //�� �����޾�

            pXLColumn[22] = 15;     //�����޾�
            pXLColumn[23] = 34;     //�Ѱ�����
            pXLColumn[24] = 25;     //�� �����޾�

            pXLColumn[25] = 4;      //���
        }

        private void SetArray6_BHK(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[27];
            pXLColumn = new int[26];

            pGDColumn[0] = pGrid.GetColumnToIndex("WAGE_TYPE_NAME");            //Title
            pGDColumn[1] = pGrid.GetColumnToIndex("DEPT_NAME");                 //�μ�
            pGDColumn[2] = pGrid.GetColumnToIndex("POST_NAME");                 //����
            pGDColumn[3] = pGrid.GetColumnToIndex("PERSON_NUM");                //���
            pGDColumn[4] = pGrid.GetColumnToIndex("NAME");                      //�̸�

            pGDColumn[5] = pGrid.GetColumnToIndex("NAME");                      //����
            pGDColumn[6] = pGrid.GetColumnToIndex("PERSON_NUM");                //���
            pGDColumn[7] = pGrid.GetColumnToIndex("DEPT_NAME");                 //�μ�
            pGDColumn[8] = pGrid.GetColumnToIndex("POST_NAME");                 //����
            pGDColumn[9] = pGrid.GetColumnToIndex("JOB_CLASS_NAME");            //����
            pGDColumn[10] = pGrid.GetColumnToIndex("SUPPLY_DATE");              //������
            pGDColumn[11] = pGrid.GetColumnToIndex("BANK_NAME");                //�Ա�����
            pGDColumn[12] = pGrid.GetColumnToIndex("BANK_ACCOUNTS");            //�Աݰ���

            pGDColumn[13] = pGrid.GetColumnToIndex("BASIC_AMOUNT");             //�⺻��
            pGDColumn[14] = pGrid.GetColumnToIndex("HOURLY_AMOUNT");            //�ñ�
            pGDColumn[15] = pGrid.GetColumnToIndex("GENERAL_HOURLY_AMOUNT");    //���ñ�

            pGDColumn[16] = pGrid.GetColumnToIndex("TOT_PAY_DED_AMOUNT");       //�޿� �Ѱ�����
            pGDColumn[17] = pGrid.GetColumnToIndex("TOT_PAY_SUP_AMOUNT");       //�޿� �����޾�
            pGDColumn[18] = pGrid.GetColumnToIndex("REAL_PAY_AMOUNT");          //�޿� �����޾�

            pGDColumn[19] = pGrid.GetColumnToIndex("TOT_BONUS_DED_AMOUNT");     //�� �Ѱ�����
            pGDColumn[20] = pGrid.GetColumnToIndex("TOT_BONUS_SUP_AMOUNT");     //�� �����޾�
            pGDColumn[21] = pGrid.GetColumnToIndex("REAL_BONUS_AMOUNT");        //�� �����޾�

            pGDColumn[22] = pGrid.GetColumnToIndex("TOT_SUPPLY_AMOUNT");        //�����޾�
            pGDColumn[23] = pGrid.GetColumnToIndex("TOT_DED_AMOUNT");           //�Ѱ�����
            pGDColumn[24] = pGrid.GetColumnToIndex("REAL_AMOUNT");              //�� �����޾�

            pGDColumn[25] = pGrid.GetColumnToIndex("DESCRIPTION");              //���
            pGDColumn[26] = pGrid.GetColumnToIndex("NOTIFICATION");             //�˸�

            pXLColumn[0] = 3;       //Title
            pXLColumn[1] = 7;       //�μ�
            pXLColumn[2] = 7;       //����
            pXLColumn[3] = 7;       //���
            pXLColumn[4] = 7;       //�̸�

            pXLColumn[5] = 8;       //����
            pXLColumn[6] = 21;      //���
            pXLColumn[7] = 35;      //�μ�
            pXLColumn[8] = 8;       //����
            pXLColumn[9] = 21;      //����
            pXLColumn[10] = 35;     //������
            pXLColumn[11] = 8;      //�Ա�����
            pXLColumn[12] = 21;     //�Աݰ���

            pXLColumn[13] = 3;     //�⺻�ϱ�
            pXLColumn[14] = 0;     //�ñ�
            pXLColumn[15] = 11;     //���ñ�

            pXLColumn[16] = 33;     //�޿� �Ѱ�����
            pXLColumn[17] = 12;     //�޿� �����޾�
            pXLColumn[18] = 33;     //�޿� �����޾�

            pXLColumn[19] = 33;     //�� �Ѱ�����
            pXLColumn[20] = 12;     //�� �����޾�
            pXLColumn[21] = 33;     //�� �����޾�

            pXLColumn[22] = 12;     //�����޾�
            pXLColumn[23] = 33;     //�Ѱ�����
            pXLColumn[24] = 26;     //�� �����޾�

            pXLColumn[25] = 3;      //���
        }

        #endregion;

        #region ----- Convert String Method ----

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
            catch (System.Exception ex)
            {
                mAppInterface.OnAppMessageEvent(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
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

        #region ----- Line Write Method : Header �� �������� -----

        private void XLLine_HEADER(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGridCol)
        {
            int vXLine = pXLine + 1; //������ ������ ǥ�õǴ� �� ��ȣ

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;

            bool IsConvert = false;
            try
            {
                mPrinting.XLActiveSheet(mTargetSheet);

                //Title
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[0]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine-1, 1, vConvertString);

                //����
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[4]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}({1})", vConvertString, pGrid.GetCellValue(pGridRow, pGridCol[3]));
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 6, vConvertString);

                //�μ���
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[1]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 6, vConvertString);

                //����
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[2]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 6, vConvertString);

                

                

                //�����޾�
                //-------------------------------------------------------------------
                vXLine = vXLine + 9 ;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[18]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //�Ѱ����� 
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[19]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 42, vConvertString);

                //���ñ�
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[11]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);
                //��_�����޾�
                //-------------------------------------------------------------------
                //vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[20]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 42, vConvertString);

                //�⺻�� 
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[9]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //��� 
                //-------------------------------------------------------------------
                //vXLine = vXLine ;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[21]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("[{0}]", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine , 1, vConvertString);
                ////-------------------------------------------------------------------
                //vXLine = vXLine + 3;
                ////-------------------------------------------------------------------

                ////���
                //vObject = pGrid.GetCellValue(pGridRow, pGridCol[21]);
                //IsConvert = IsConvertString(vObject, out vConvertString);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0}", vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //ȸ���
                //-------------------------------------------------------------------
                //vXLine = vXLine ;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[22]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("[{0}]", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine -8, 2, vConvertString);

                

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLLine_HEADER_SIK(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGridCol)
        {
            int vXLine = pXLine + 1; //������ ������ ǥ�õǴ� �� ��ȣ

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;

            bool IsConvert = false;
            try
            {
                mPrinting.XLActiveSheet(mTargetSheet);

                //Title
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[0]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 2, vConvertString);

                //����
                //-------------------------------------------------------------------
                vXLine = vXLine + 5;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[4]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}({1})", vConvertString, pGrid.GetCellValue(pGridRow, pGridCol[3]));
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 14, vConvertString);

                //�μ���
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[1]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 14, vConvertString);

                //����
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[2]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 14, vConvertString);

                //���ñ�
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[11]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 14, vConvertString);

                //�⺻�� 
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[9]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 32, vConvertString);

                //�����޾�
                //-------------------------------------------------------------------
                vXLine = vXLine + 39;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[18]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 15, vConvertString);

                //�Ѱ����� 
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[19]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 35, vConvertString);

                //��_�����޾�
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[20]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 22, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                //���
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[21]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[22]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("[{0}]", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 33, vConvertString);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }
        
        private void XLLine_HEADER_SIV(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGridCol)
        {
            int vXLine = pXLine + 1; //������ ������ ǥ�õǴ� �� ��ȣ

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;

            bool IsConvert = false;
            try
            {
                mPrinting.XLActiveSheet(mTargetSheet);

                //Title
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[0]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLCellMerge(vXLine, 2, 4, 35, true);
                mPrinting.XLSetCell(vXLine, 2, vConvertString);

                //����
                //-------------------------------------------------------------------
                vXLine = vXLine + 5;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[4]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}({1})", vConvertString, pGrid.GetCellValue(pGridRow, pGridCol[3]));
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //�μ���
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[1]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //����
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[2]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //���ñ�
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[11]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 12, vConvertString);

                //�⺻�� 
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[9]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 29, vConvertString);

                //�����޾�
                //-------------------------------------------------------------------
                vXLine = vXLine + 35;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[18]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 13, vConvertString);

                //�Ѱ����� 
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[19]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 32, vConvertString);

                //��_�����޾�
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[20]);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,##0}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 20, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                //���
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[21]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 1, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------
                vObject = pGrid.GetCellValue(pGridRow, pGridCol[22]);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("[{0}]", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, 30, vConvertString);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }


        #endregion;

        #region ----- Line Write Method : Adapter ���� �μ�(����ٹ�) -----

        private void XLLine_OT(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start+12;    //���� �μ�� Row ���� ��ġ
            int vXL_Col = 1;
            int vMax_Row_Count = 10;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //����ٹ��׸��
                    vObject = vRow["OT_TYPE_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, vXL_Col, vConvertString);

                    //����ٹ��ð�
                    vObject = vRow["OT_TIME"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, vXL_Col+6, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    if(vXL_Row == pXL_Row_Start + 17)
                    {
                        vXL_Row = pXL_Row_Start + 12;
                        vXL_Col = 11;
                    }
                    //-------------------------------------------------------------------
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLLine_OT_SIK(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 13;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //����ٹ��׸��
                    vObject = vRow["OT_TYPE_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 1, vConvertString);

                    //����ٹ��ð�
                    vObject = vRow["OT_TIME"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 15, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLLine_OT_SIV(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 10;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //����ٹ��׸��
                    vObject = vRow["OT_TYPE_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 1, vConvertString);

                    //����ٹ��ð�
                    vObject = vRow["OT_TIME"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###.##}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 13, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
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

        #region ----- Line Write Method : Adapter ���� �μ�(���°�) -----

        private void XLLine_DUTY_SIK(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 10;        //�ִ� �μ� Row��
            //int vMax_Row_Count2 = 5;        //�ִ� �μ� Row��
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //�׸��
                    vObject = vRow["DUTY_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 22, vConvertString);

                    //Ƚ��
                    vObject = vRow["DUTY_COUNT"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 35, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLLine_DUTY_SIV(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 10;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //�׸��
                    vObject = vRow["DUTY_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 20, vConvertString);

                    //Ƚ��
                    vObject = vRow["DUTY_COUNT"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 32, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
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

        #region ----- Line Write Method : Adapter ���� �μ�(�����׸�) -----

        private void XLLine_ALLOWANCE(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 15;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //�׸��
                    vObject = vRow["ALLOWANCE_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row +1, 21, vConvertString);

                    //���ޱݾ�
                    vObject = vRow["ALLOWANCE_AMOUNT"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row+1, 29, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLLine_ALLOWANCE_SIK(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 18;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //�׸��
                    vObject = vRow["ALLOWANCE_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 1, vConvertString);

                    //���ޱݾ�
                    vObject = vRow["ALLOWANCE_AMOUNT"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 15, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLLine_ALLOWANCE_SIV(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 15;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //�׸��
                    vObject = vRow["ALLOWANCE_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 1, vConvertString);

                    //���ޱݾ�
                    vObject = vRow["ALLOWANCE_AMOUNT"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 13, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
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

        #region ----- Line Write Method : Adapter ���� �μ�(�����׸�) -----

        private void XLLine_DEDUCTION(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 18;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //�׸��
                    vObject = vRow["DEDUCTION_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row+1, 34, vConvertString);

                    //���� �ݾ�
                    vObject = vRow["DEDUCTION_AMOUNT"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row+1, 42, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLLine_DEDUCTION_SIK(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 18;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //�׸��
                    vObject = vRow["DEDUCTION_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 22, vConvertString);

                    //���� �ݾ�
                    vObject = vRow["DEDUCTION_AMOUNT"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 35, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLLine_DEDUCTION_SIV(InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter, int pXL_Row_Start)
        {
            int vXL_Row = pXL_Row_Start;    //���� �μ�� Row ���� ��ġ
            int vMax_Row_Count = 15;        //�ִ� �μ� Row��

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            mPrinting.XLActiveSheet(mTargetSheet);
            try
            {
                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    if (vMax_Row_Count < 0)
                    {
                        return;
                    }

                    //�׸��
                    vObject = vRow["DEDUCTION_NAME"];
                    IsConvert = IsConvertString(vObject, out vConvertString);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0}", vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 20, vConvertString);

                    //���� �ݾ�
                    vObject = vRow["DEDUCTION_AMOUNT"];
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:###,###,###,###,###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXL_Row, 32, vConvertString);

                    //-------------------------------------------------------------------
                    vXL_Row = vXL_Row + 1;
                    vMax_Row_Count = vMax_Row_Count - 1;
                    //-------------------------------------------------------------------
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


        #region ----- Line Write 1 Method -----

        //�޿� ����
        private int XLLine_1(System.Data.DataRow pRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vDBColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //�޿� ���޸�
                vDBColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pRow[vDBColumnIndex];
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

                //�޿� ���ޱݾ�
                vDBColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
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

        #region ----- Line Write 2 Method -----

        //�޿� ����
        private int XLLine_2(System.Data.DataRow pRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vDBColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //�޿� ������
                vDBColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pRow[vDBColumnIndex];
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

                //�޿� �����ݾ�
                vDBColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
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

        #region ----- Line Write 3 Method -----

        //�� ����
        private int XLLine_3(System.Data.DataRow pRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vDBColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //�� ���޸�
                vDBColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pRow[vDBColumnIndex];
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

                //�� ���ޱݾ�
                vDBColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
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

        #region ----- Line Write 4 Method -----

        //�� ����
        private int XLLine_4(System.Data.DataRow pRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vDBColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //�� ������
                vDBColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pRow[vDBColumnIndex];
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

                //�� �����ݾ�
                vDBColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
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

        #region ----- Line Write 5 Method -----

        //�ٹ��ð� �� �ΰ�����
        private int XLLine_5(System.Data.DataRow pRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vDBColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //����(����)
                vDBColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�߰�(����)
                vDBColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //���°���(����)
                vDBColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //������-�ٹ�
                vDBColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������-����
                vDBColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������-�߰�
                vDBColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //������-�ٹ�
                vDBColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������-����
                vDBColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������-�߰�
                vDBColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                //�ٹ�(�ΰ�����)
                vDBColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����(�ΰ�����)
                vDBColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����(�ΰ�����)
                vDBColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����(�ΰ�����)
                vDBColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����(�ΰ�����)
                vDBColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�̱ٹ�(�ΰ�����)
                vDBColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������(�ΰ�����)
                vDBColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
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


        //�ٹ��ð� �� �ΰ�����
        private int XLLine_5_BHK(System.Data.DataRow pRow, int pXLine, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vDBColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet(mTargetSheet);

                //�⺻ �ٹ��ð�
                vDBColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����(����)
                vDBColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�߰�
                vDBColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������-�ٹ�
                vDBColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //������-����
                vDBColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //���°���
                vDBColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    if (vConvertDecimal != 0)
                    {
                        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                ////-------------------------------------------------------------------
                //vXLine = vXLine + 1;
                ////-------------------------------------------------------------------

                ////������-�ٹ�
                //vDBColumnIndex = pGDColumn[6];
                //vXLColumnIndex = pXLColumn[6];
                //vObject = pRow[vDBColumnIndex];
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    if (vConvertDecimal != 0)
                //    {
                //        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                //    }
                //    else
                //    {
                //        vConvertString = string.Empty;
                //    }
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                ////������-����
                //vDBColumnIndex = pGDColumn[7];
                //vXLColumnIndex = pXLColumn[7];
                //vObject = pRow[vDBColumnIndex];
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    if (vConvertDecimal != 0)
                //    {
                //        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                //    }
                //    else
                //    {
                //        vConvertString = string.Empty;
                //    }
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                ////������-�߰�
                //vDBColumnIndex = pGDColumn[8];
                //vXLColumnIndex = pXLColumn[8];
                //vObject = pRow[vDBColumnIndex];
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    if (vConvertDecimal != 0)
                //    {
                //        vConvertString = string.Format("{0:#,##0.###}", vConvertDecimal);
                //    }
                //    else
                //    {
                //        vConvertString = string.Empty;
                //    }
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                //-------------------------------------------------------------------
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                //�ٹ��ϼ�(�ΰ�����)
                vDBColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�����ϼ�
                vDBColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                ////�������ϼ�
                //vDBColumnIndex = pGDColumn[11];
                //vXLColumnIndex = pXLColumn[11];
                //vObject = pRow[vDBColumnIndex];
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                //�����ϼ�(�ΰ�����)
                vDBColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //���(�ΰ�����)
                vDBColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //����/����/����(�ΰ�����)
                vDBColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pRow[vDBColumnIndex];
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                ////�����뿬��(�ΰ�����)
                //vDBColumnIndex = pGDColumn[15];
                //vXLColumnIndex = pXLColumn[15];
                //vObject = pRow[vDBColumnIndex];
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////������뿬��(�ΰ�����)
                //vDBColumnIndex = pGDColumn[16];
                //vXLColumnIndex = pXLColumn[16];
                //vObject = pRow[vDBColumnIndex];
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
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

        #region ----- Line Write 6 Method -----

        //Heaer �� ��������, �ѱݾ�
        private int XLLine_6(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pCourse)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //Title
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 10;
                //-------------------------------------------------------------------

                //�μ�
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                //����
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                //���
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
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

                //�̸�
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 9;
                //-------------------------------------------------------------------

                //����
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
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

                //�μ�
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //����
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
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

                //����
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

                //������
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //�Ա�����
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

                //�Աݰ���
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
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
                vXLine = vXLine + 10;
                //-------------------------------------------------------------------

                //�⺻��
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�ñ�
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //���ñ�
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                if (pCourse == "DAILY")
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 16;
                    //-------------------------------------------------------------------

                    //�޿�_�Ѱ�����
                    vGDColumnIndex = pGDColumn[16];
                    vXLColumnIndex = pXLColumn[16];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    //�޿�_�����޾�
                    vGDColumnIndex = pGDColumn[17];
                    vXLColumnIndex = pXLColumn[17];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //�޿�_�����޾�
                    vGDColumnIndex = pGDColumn[18];
                    vXLColumnIndex = pXLColumn[18];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 6;
                    //-------------------------------------------------------------------

                    //��_�Ѱ�����
                    vGDColumnIndex = pGDColumn[19];
                    vXLColumnIndex = pXLColumn[19];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------

                    //��_�����޾�
                    vGDColumnIndex = pGDColumn[20];
                    vXLColumnIndex = pXLColumn[20];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //��_�����޾�
                    vGDColumnIndex = pGDColumn[21];
                    vXLColumnIndex = pXLColumn[21];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 1;
                    //-------------------------------------------------------------------
                }
                else if (pCourse == "MONTH")
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 25;
                    //-------------------------------------------------------------------
                }

                //�����޾�
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�Ѱ�����
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //��_�����޾�
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                //���
                vGDColumnIndex = pGDColumn[26];  // �˸� //
                vXLColumnIndex = pXLColumn[25];                
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
                vXLine = vXLine + 1;
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


        //Heaer �� ��������, �ѱݾ�
        private int XLLine_6_BHK(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pCourse)
        {
            int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet(mTargetSheet);

                //Title
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                //�μ�
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                //����
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                //���
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
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

                //�̸�
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

                //-------------------------------------------------------------------
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                //����
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
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

                //�μ�
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
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
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //����
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
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

                //����
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

                //������
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //�Ա�����
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

                //�Աݰ���
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
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
                vXLine = vXLine + 8;
                //-------------------------------------------------------------------

                //�⺻��
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                ////�ñ�
                //vGDColumnIndex = pGDColumn[14];
                //vXLColumnIndex = pXLColumn[14];
                //vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                //���ñ�
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                if (pCourse == "DAILY")
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 28;
                    //-------------------------------------------------------------------

                    //�޿�_�Ѱ�����
                    vGDColumnIndex = pGDColumn[16];
                    vXLColumnIndex = pXLColumn[16];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
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

                    //�޿�_�����޾�
                    vGDColumnIndex = pGDColumn[17];
                    vXLColumnIndex = pXLColumn[17];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //�޿�_�����޾�
                    vGDColumnIndex = pGDColumn[18];
                    vXLColumnIndex = pXLColumn[18];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //-------------------------------------------------------------------
                    vXLine = vXLine + 10;
                    //-------------------------------------------------------------------

                    //��_�Ѱ�����
                    vGDColumnIndex = pGDColumn[19];
                    vXLColumnIndex = pXLColumn[19];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
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

                    //��_�����޾�
                    vGDColumnIndex = pGDColumn[20];
                    vXLColumnIndex = pXLColumn[20];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    //��_�����޾�
                    vGDColumnIndex = pGDColumn[21];
                    vXLColumnIndex = pXLColumn[21];
                    vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:#}", vConvertDecimal);
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
                else if (pCourse == "MONTH")
                {
                    //-------------------------------------------------------------------
                    vXLine = vXLine + 44;
                    //-------------------------------------------------------------------
                }

                //�����޾�
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�Ѱ�����
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                //��_�����޾�
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                //���
                vGDColumnIndex = pGDColumn[26];  // �˸� //
                vXLColumnIndex = pXLColumn[25];
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

        //30�徿
        #region ----- Excel Main Wirte  Method ----

        public int WriteMain(string pCourse, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_MONTH_PAYMENT, InfoSummit.Win.ControlAdv.ISDataAdapter pData_PAY_ALLOWANCE, InfoSummit.Win.ControlAdv.ISDataAdapter pData_PAY_DEDUCTION, InfoSummit.Win.ControlAdv.ISDataAdapter pData_DUTY_INFO, InfoSummit.Win.ControlAdv.ISDataAdapter pData_BONUS_ALLOWANCE, InfoSummit.Win.ControlAdv.ISDataAdapter pData_BONUS_DEDUCTION, string pSeal_Stamp)
        {
            string vMessageText = string.Empty;
            object vObject = null;
            string vBoxCheck = string.Empty;
            string vWAGE_TYPE = string.Empty;
            string vPAY_TYPE = string.Empty;

            int vIndexWAGE_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIndexPAY_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");
            int vIndexPRINT_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("PRINT_TYPE");

            int vIndexCheckBox = pGrid_MONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN");
            string vCheckedString = pGrid_MONTH_PAYMENT.GridAdvExColElement[vIndexCheckBox].CheckedString;

            bool isOpen = XLFileOpen();

            mPageNumber = 0;

            int[] vGDColumn_1;
            int[] vXLColumn_1;

            int[] vGDColumn_2;
            int[] vXLColumn_2;

            int[] vGDColumn_3;
            int[] vXLColumn_3;

            int[] vGDColumn_4;
            int[] vXLColumn_4;

            int[] vGDColumn_5;
            int[] vXLColumn_5;

            int[] vGDColumn_6;
            int[] vXLColumn_6;

            int vTotalRow = pGrid_MONTH_PAYMENT.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 0;

            int vSecondPrinting = 29; //30��°�� �μ�
            int vCountPrinting = 0;

            SetArray1(pData_PAY_ALLOWANCE.OraSelectData, out vGDColumn_1, out vXLColumn_1);
            SetArray2(pData_PAY_DEDUCTION.OraSelectData, out vGDColumn_2, out vXLColumn_2);
            SetArray3(pData_BONUS_ALLOWANCE.OraSelectData, out vGDColumn_3, out vXLColumn_3);
            SetArray4(pData_BONUS_DEDUCTION.OraSelectData, out vGDColumn_4, out vXLColumn_4);
            SetArray5(pData_DUTY_INFO.OraSelectData, out vGDColumn_5, out vXLColumn_5);
            SetArray6(pGrid_MONTH_PAYMENT, out vGDColumn_6, out vXLColumn_6);

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;                
                 
                vMessageText = string.Format("Grid : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIndexCheckBox);
                vBoxCheck = ConvertString(vObject);
                if (ConvertString(pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIndexPRINT_TYPE)) == "1")
                {
                    if (vBoxCheck == vCheckedString)
                    {//üũ�� ����߿� �μ���Ǹ� �μ�//
                        pGrid_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIndexCheckBox);
                        pGrid_MONTH_PAYMENT.Focus();
                        pGrid_MONTH_PAYMENT.CurrentCellActivate(vRow, vIndexCheckBox);
                        if (isOpen == true)
                        {
                            vCountPrinting++;

                            vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIndexWAGE_TYPE);
                            vWAGE_TYPE = ConvertString(vObject);
                            vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIndexPAY_TYPE);
                            vPAY_TYPE = ConvertString(vObject);

                            if (vWAGE_TYPE == "P1" && (vPAY_TYPE == "2" || vPAY_TYPE == "4"))
                            {
                                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "DAILY", pSeal_Stamp);
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                                //������
                                int vLinePrinting_1 = vPrintingLine + 41;
                                for (int vRow1 = 0; vRow1 < pData_PAY_ALLOWANCE.OraSelectData.Rows.Count; vRow1++)
                                {
                                    vLinePrinting_1 = XLLine_1(pData_PAY_ALLOWANCE.OraSelectData.Rows[vRow1], vLinePrinting_1, vGDColumn_1, vXLColumn_1); //�޿� ����
                                }

                                int vLinePrinting_2 = vPrintingLine + 41;
                                for (int vRow2 = 0; vRow2 < pData_PAY_DEDUCTION.OraSelectData.Rows.Count; vRow2++)
                                {
                                    vLinePrinting_2 = XLLine_2(pData_PAY_DEDUCTION.OraSelectData.Rows[vRow2], vLinePrinting_2, vGDColumn_2, vXLColumn_2); //�޿� ����
                                }

                                int vLinePrinting_3 = vPrintingLine + 55;
                                for (int vRow3 = 0; vRow3 < pData_BONUS_ALLOWANCE.OraSelectData.Rows.Count; vRow3++)
                                {
                                    vLinePrinting_3 = XLLine_3(pData_BONUS_ALLOWANCE.OraSelectData.Rows[vRow3], vLinePrinting_3, vGDColumn_3, vXLColumn_3); //�� ����
                                }

                                int vLinePrinting_4 = vPrintingLine + 55;
                                for (int vRow4 = 0; vRow4 < pData_BONUS_DEDUCTION.OraSelectData.Rows.Count; vRow4++)
                                {
                                    vLinePrinting_4 = XLLine_4(pData_BONUS_DEDUCTION.OraSelectData.Rows[vRow4], vLinePrinting_4, vGDColumn_4, vXLColumn_4); //�� ����
                                }

                                int vLinePrinting_5 = vPrintingLine + 31;
                                for (int vRow5 = 0; vRow5 < pData_DUTY_INFO.OraSelectData.Rows.Count; vRow5++)
                                {
                                    vLinePrinting_5 = XLLine_5(pData_DUTY_INFO.OraSelectData.Rows[vRow5], vLinePrinting_5, vGDColumn_5, vXLColumn_5); //�ٹ��ð� �� �ΰ�����
                                }

                                vPrintingLine = XLLine_6(pGrid_MONTH_PAYMENT, vRow, vPrintingLine, vGDColumn_6, vXLColumn_6, "DAILY"); //Heaer �� ��������, �ѱݾ�
                            }
                            else
                            {
                                mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "MONTH", pSeal_Stamp);
                                vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                                //������
                                int vLinePrinting_1 = vPrintingLine + 41;
                                for (int vRow1 = 0; vRow1 < pData_PAY_ALLOWANCE.OraSelectData.Rows.Count; vRow1++)
                                {
                                    vLinePrinting_1 = XLLine_1(pData_PAY_ALLOWANCE.OraSelectData.Rows[vRow1], vLinePrinting_1, vGDColumn_1, vXLColumn_1); //�޿� ����
                                }

                                int vLinePrinting_2 = vPrintingLine + 41;
                                for (int vRow2 = 0; vRow2 < pData_PAY_DEDUCTION.OraSelectData.Rows.Count; vRow2++)
                                {
                                    vLinePrinting_2 = XLLine_2(pData_PAY_DEDUCTION.OraSelectData.Rows[vRow2], vLinePrinting_2, vGDColumn_2, vXLColumn_2); //�޿� ����
                                }

                                int vLinePrinting_5 = vPrintingLine + 31;
                                for (int vRow5 = 0; vRow5 < pData_DUTY_INFO.OraSelectData.Rows.Count; vRow5++)
                                {
                                    vLinePrinting_5 = XLLine_5(pData_DUTY_INFO.OraSelectData.Rows[vRow5], vLinePrinting_5, vGDColumn_5, vXLColumn_5); //�ٹ��ð� �� �ΰ�����
                                }

                                vPrintingLine = XLLine_6(pGrid_MONTH_PAYMENT, vRow, vPrintingLine, vGDColumn_6, vXLColumn_6, "MONTH"); //Heaer �� ��������, �ѱݾ�
                            }

                            if (pCourse == "PRINT")
                            {
                                if (vTotalRow == vRowCount)
                                {
                                    Printing(1, vCountPrinting);
                                }
                                else if (vSecondPrinting < vCountPrinting)
                                {
                                    Printing(1, vCountPrinting);

                                    mPrinting.XLOpenFileClose();
                                    isOpen = XLFileOpen();

                                    vCountPrinting = 0;
                                    vPrintingLine = 1;
                                    mCopyLineSUM = 1;
                                }
                            }
                            else if (pCourse == "FILE")
                            {
                                if (vTotalRow == vRowCount)
                                {
                                    SAVE("PAY_");
                                }
                                else if (vSecondPrinting < vCountPrinting)
                                {
                                    SAVE("PAY_");

                                    mPrinting.XLOpenFileClose();
                                    isOpen = XLFileOpen();

                                    vCountPrinting = 0;
                                    vPrintingLine = 1;
                                    mCopyLineSUM = 1;
                                }
                            }
                            pGrid_MONTH_PAYMENT.SetCellValue(vRow, vIndexCheckBox, "N");
                        }
                    }
                    else if (vTotalRow == vRowCount)
                    {
                        if (isOpen == true)
                        {
                            if (pCourse == "PRINT")
                            {
                                Printing(1, vCountPrinting);
                            }
                            else if (pCourse == "FILE")
                            {
                                SAVE("PAY_");
                            }
                        }
                    }
                }
            }
            return mPageNumber;
        }

         
        public int WriteMain(string pOUTPUT_TYPE, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_MONTH_PAYMENT
                                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PAY_ALLOWANCE
                                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PAY_DEDUCTION
                                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_MONTH_DUTY
                                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_MONTH_OT
                                            , string pSeal_Stamp
            )
        {
            string vMessageText = string.Empty;
            object vObject = null;
            string vWAGE_TYPE = string.Empty;
            string vPAY_TYPE = string.Empty;
            string vCheckedString = "N";
            string vCheckedString2 = "N";

            int vIDX_CheckBox = pGrid_MONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN");
            int vIDX_WAGE_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIDX_PAY_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");

            int[] vGridCol;

            bool isOpen = XLFileOpen();

            mPageNumber = 0;
            mCopyLineSUM = 0;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 46;
            mCopy_EndRow = 20;

            mCurrentRow = 1;
            mCurrentCol = 1;

            int vTotalRow = pGrid_MONTH_PAYMENT.RowCount;
            int vRowCount = 0;

            SetArray1(pGrid_MONTH_PAYMENT, out vGridCol);

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;

                vMessageText = string.Format("Printing Rows : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);

                System.Windows.Forms.Application.UseWaitCursor = true;
                pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                System.Windows.Forms.Application.DoEvents();

                vCheckedString = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_CheckBox).ToString();
                if (vCheckedString == "Y")
                {
                    if (isOpen == true)
                    {
                        pGrid_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIDX_CheckBox);

                        mPrintPage++;

                        vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_WAGE_TYPE);
                        vWAGE_TYPE = ConvertString(vObject);
                        vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_PAY_TYPE);
                        vPAY_TYPE = ConvertString(vObject);

                        // j = 1;
                        if(mPageNumber >0)
                        {
                            if(mPageNumber > 3 &&  mPageNumber%3 == 0)
                            {
                                mCurrentRow = mCurrentRow + mCopy_EndRow -1;
                              //  mCopy_EndRow = 19;
                            }
                            else
                            {
                                mCurrentRow = mCurrentRow + mCopy_EndRow;
                               // mCopy_EndRow = 20;
                            }
                        }
                            
                            //(mCopy_EndRow * mPageNumber)+1;  //���� �μ� row ��ġ ���� : �μ�row�� �������� + 1 �� ������ ������ ���.
                    
                        
                        mCopyLineSUM = CopyAndPaste(mPrinting, mCurrentRow, mSourceSheet1, pSeal_Stamp);

                        //vRow++;
                        // �������� �μ�
                        XLLine_HEADER(pGrid_MONTH_PAYMENT, vRow, mCurrentRow, vGridCol); //Heaer �� ��������, �ѱݾ�

                        //����ٹ� �μ�
                        XLLine_OT(pIDA_MONTH_OT, mCurrentRow);

                        //���°� �μ�
                        //XLLine_DUTY(pIDA_MONTH_DUTY, mCurrentRow + 22);

                        //�����׸� �μ�
                        XLLine_ALLOWANCE(pIDA_PAY_ALLOWANCE, mCurrentRow);

                        //�����׸� �μ�
                        XLLine_DEDUCTION(pIDA_PAY_DEDUCTION, mCurrentRow);

                        vCheckedString2 = pGrid_MONTH_PAYMENT.GetCellValue(vRow + 1, vIDX_CheckBox).ToString();


                        ////���忡 ���� �μ�////
                        ////VCC�� A5�� �μ��ϱ� ������ �ּ�//
                        //if (vCheckedString2 == "Y")
                        //{
                        //    for (int vRow2 = 0; vRow2 < 1; vRow2++)
                        //    {
                        //        vRow++;
                        //        vRowCount++;

                        //        pGrid_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIDX_CheckBox);
                        //        pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                        //        mCurrentCol = mCopy_EndCol + 1;  //���� �μ� row ��ġ ���� : �μ�row�� �������� + 1 �� ������ ������ ���.

                        //        mCopyLineSUM = CopyAndPaste2(mPrinting, mCurrentRow, mCurrentCol, mSourceSheet1);


                        //        // �������� �μ�
                        //        XLLine_HEADER2(pGrid_MONTH_PAYMENT, vRow, mCurrentRow, vGridCol); //Heaer �� ��������, �ѱݾ�

                        //        //����ٹ� �μ�
                        //        XLLine_OT2(pIDA_MONTH_OT, mCurrentRow + 16);

                        //        //���°� �μ�
                        //        XLLine_DUTY2(pIDA_MONTH_DUTY, mCurrentRow + 16);

                        //        //�����׸� �μ�
                        //        XLLine_ALLOWANCE2(pIDA_PAY_ALLOWANCE, mCurrentRow + 37);

                        //        //�����׸� �μ�
                        //        XLLine_DEDUCTION2(pIDA_PAY_DEDUCTION, mCurrentRow + 37);

                        //        //if (vTotalRow == vRowCount)
                        //        //{
                        //        //    if (pCourse == "PRINT")
                        //        //    {
                        //        //        Printing(1, mPageNumber);
                        //        //    }
                        //        //    else if (pCourse == "FILE")
                        //        //    {
                        //        //        SAVE("PAY_");
                        //        //    }
                        //        //}
                        //        //else if (mMaxPrintPage < mPageNumber)
                        //        //{
                        //        //    if (pCourse == "PRINT")
                        //        //    {
                        //        //        Printing(1, mPageNumber);

                        //        //        mPrinting.XLOpenFileClose();
                        //        //        isOpen = XLFileOpen();
                        //        //    }
                        //        //    else if (pCourse == "FILE")
                        //        //    {
                        //        //        SAVE("PAY_");

                        //        //        mPrinting.XLOpenFileClose();
                        //        //        isOpen = XLFileOpen();
                        //        //    }
                        //        //    mPageNumber = 0;
                        //        //    mCurrentRow = 1;
                        //        //    mCopyLineSUM = 1;
                        //        //}





                        //    }

                        //}
                        if (vTotalRow == vRowCount)
                        {
                            if (pOUTPUT_TYPE == "PRINT")
                            {
                                Printing(1, mPageNumber);
                            }
                            else if (pOUTPUT_TYPE == "FILE")
                            {
                                SAVE("PAY_");
                            }
                        }
                        else if (mMaxPrintPage < mPageNumber)
                        {
                            if (pOUTPUT_TYPE == "PRINT")
                            {
                                Printing(1, mPageNumber);

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();
                            }
                            else if (pOUTPUT_TYPE == "FILE")
                            {
                                SAVE("PAY_");

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();
                            }
                            mPageNumber = 0;
                            mCurrentRow = 1;
                            mCopyLineSUM = 1;
                        }
                    }


                    pGrid_MONTH_PAYMENT.SetCellValue(vRow, vIDX_CheckBox, "N");
                }
                else if (vTotalRow == vRowCount)
                {
                    if (isOpen == true)
                    {
                        if (pOUTPUT_TYPE == "PRINT")
                        {
                            Printing(1, mPageNumber);
                        }
                        else if (pOUTPUT_TYPE == "FILE")
                        {
                            SAVE("PAY_");
                        }
                    }
                }
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();

            return mPageNumber;
        }

        public int WriteMain_SIK(string pOUTPUT_TYPE, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_MONTH_PAYMENT
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PAY_ALLOWANCE
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PAY_DEDUCTION
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_MONTH_DUTY
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_MONTH_OT
                                                   , string pSeal_Stamp
            )
        {
            string vMessageText = string.Empty;
            object vObject = null;
            string vWAGE_TYPE = string.Empty;
            string vPAY_TYPE = string.Empty;
            string vCheckedString = "N";
            string vCheckedString2 = "N";

            int vIDX_CheckBox = pGrid_MONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN");
            int vIDX_WAGE_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIDX_PAY_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");

            int[] vGridCol;

            bool isOpen = XLFileOpen();

            mPageNumber = 0;
            mCopyLineSUM = 0;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 45;
            mCopy_EndRow = 59;

            mCurrentRow = 1;
            mCurrentCol = 1;

            int vTotalRow = pGrid_MONTH_PAYMENT.RowCount;
            int vRowCount = 0;

            SetArray1(pGrid_MONTH_PAYMENT, out vGridCol);

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;

                vMessageText = string.Format("Printing Rows : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);

                System.Windows.Forms.Application.UseWaitCursor = true;
                pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                System.Windows.Forms.Application.DoEvents();

                vCheckedString = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_CheckBox).ToString();
                if (vCheckedString == "Y")
                {
                    if (isOpen == true)
                    {
                        pGrid_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIDX_CheckBox);

                        mPrintPage++;

                        vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_WAGE_TYPE);
                        vWAGE_TYPE = ConvertString(vObject);
                        vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_PAY_TYPE);
                        vPAY_TYPE = ConvertString(vObject);

                        mCurrentRow = (mCopy_EndRow * mPageNumber) + 1;  //���� �μ� row ��ġ ���� : �μ�row�� �������� + 1 �� ������ ������ ���.

                        mCopyLineSUM = CopyAndPaste(mPrinting, mCurrentRow, mSourceSheet1, pSeal_Stamp);

                        //vRow++;
                        // �������� �μ�
                        XLLine_HEADER_SIK(pGrid_MONTH_PAYMENT, vRow, mCurrentRow, vGridCol); //Heaer �� ��������, �ѱݾ�

                        //����ٹ� �μ�
                        XLLine_OT_SIK(pIDA_MONTH_OT, mCurrentRow + 22);

                        //���°� �μ�
                        XLLine_DUTY_SIK(pIDA_MONTH_DUTY, mCurrentRow + 22);

                        //�����׸� �μ�
                        XLLine_ALLOWANCE_SIK(pIDA_PAY_ALLOWANCE, mCurrentRow + 34);

                        //�����׸� �μ�
                        XLLine_DEDUCTION_SIK(pIDA_PAY_DEDUCTION, mCurrentRow + 34);

                        vCheckedString2 = pGrid_MONTH_PAYMENT.GetCellValue(vRow + 1, vIDX_CheckBox).ToString();


                        ////���忡 ���� �μ�////
                        ////VCC�� A5�� �μ��ϱ� ������ �ּ�//
                        //if (vCheckedString2 == "Y")
                        //{
                        //    for (int vRow2 = 0; vRow2 < 1; vRow2++)
                        //    {
                        //        vRow++;
                        //        vRowCount++;

                        //        pGrid_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIDX_CheckBox);
                        //        pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                        //        mCurrentCol = mCopy_EndCol + 1;  //���� �μ� row ��ġ ���� : �μ�row�� �������� + 1 �� ������ ������ ���.

                        //        mCopyLineSUM = CopyAndPaste2(mPrinting, mCurrentRow, mCurrentCol, mSourceSheet1);


                        //        // �������� �μ�
                        //        XLLine_HEADER2(pGrid_MONTH_PAYMENT, vRow, mCurrentRow, vGridCol); //Heaer �� ��������, �ѱݾ�

                        //        //����ٹ� �μ�
                        //        XLLine_OT2(pIDA_MONTH_OT, mCurrentRow + 16);

                        //        //���°� �μ�
                        //        XLLine_DUTY2(pIDA_MONTH_DUTY, mCurrentRow + 16);

                        //        //�����׸� �μ�
                        //        XLLine_ALLOWANCE2(pIDA_PAY_ALLOWANCE, mCurrentRow + 37);

                        //        //�����׸� �μ�
                        //        XLLine_DEDUCTION2(pIDA_PAY_DEDUCTION, mCurrentRow + 37);

                        //        //if (vTotalRow == vRowCount)
                        //        //{
                        //        //    if (pCourse == "PRINT")
                        //        //    {
                        //        //        Printing(1, mPageNumber);
                        //        //    }
                        //        //    else if (pCourse == "FILE")
                        //        //    {
                        //        //        SAVE("PAY_");
                        //        //    }
                        //        //}
                        //        //else if (mMaxPrintPage < mPageNumber)
                        //        //{
                        //        //    if (pCourse == "PRINT")
                        //        //    {
                        //        //        Printing(1, mPageNumber);

                        //        //        mPrinting.XLOpenFileClose();
                        //        //        isOpen = XLFileOpen();
                        //        //    }
                        //        //    else if (pCourse == "FILE")
                        //        //    {
                        //        //        SAVE("PAY_");

                        //        //        mPrinting.XLOpenFileClose();
                        //        //        isOpen = XLFileOpen();
                        //        //    }
                        //        //    mPageNumber = 0;
                        //        //    mCurrentRow = 1;
                        //        //    mCopyLineSUM = 1;
                        //        //}





                        //    }

                        //}
                        if (vTotalRow == vRowCount)
                        {
                            if (pOUTPUT_TYPE == "PRINT")
                            {
                                Printing(1, mPageNumber);
                            }
                            else if (pOUTPUT_TYPE == "FILE")
                            {
                                SAVE("PAY_");
                            }
                        }
                        else if (mMaxPrintPage < mPageNumber)
                        {
                            if (pOUTPUT_TYPE == "PRINT")
                            {
                                Printing(1, mPageNumber);

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();
                            }
                            else if (pOUTPUT_TYPE == "FILE")
                            {
                                SAVE("PAY_");

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();
                            }
                            mPageNumber = 0;
                            mCurrentRow = 1;
                            mCopyLineSUM = 1;
                        }
                    }


                    pGrid_MONTH_PAYMENT.SetCellValue(vRow, vIDX_CheckBox, "N");
                }
                else if (vTotalRow == vRowCount)
                {
                    if (isOpen == true)
                    {
                        if (pOUTPUT_TYPE == "PRINT")
                        {
                            Printing(1, mPageNumber);
                        }
                        else if (pOUTPUT_TYPE == "FILE")
                        {
                            SAVE("PAY_");
                        }
                    }
                }
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();

            return mPageNumber;
        }
        public int WriteMain_SIV(string pOUTPUT_TYPE, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_MONTH_PAYMENT
                                         , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PAY_ALLOWANCE
                                         , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PAY_DEDUCTION
                                         , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_MONTH_DUTY
                                         , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_MONTH_OT
                                         , string pSeal_Stamp
            )
        {
            string vMessageText = string.Empty;
            object vObject = null;
            string vWAGE_TYPE = string.Empty;
            string vPAY_TYPE = string.Empty;
            string vCheckedString = "N";
            string vCheckedString2 = "N";

            int vIDX_CheckBox = pGrid_MONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN");
            int vIDX_WAGE_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIDX_PAY_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");

            int[] vGridCol;

            bool isOpen = XLFileOpen();

            mPageNumber = 0;
            mCopyLineSUM = 0;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 38;
            mCopy_EndRow = 54;

            mCurrentRow = 1;
            mCurrentCol = 1;

            int vTotalRow = pGrid_MONTH_PAYMENT.RowCount;
            int vRowCount = 0;

            SetArray1(pGrid_MONTH_PAYMENT, out vGridCol);

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;

                vMessageText = string.Format("Printing Rows : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);

                System.Windows.Forms.Application.UseWaitCursor = true;
                pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                System.Windows.Forms.Application.DoEvents();

                vCheckedString = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_CheckBox).ToString();
                if (vCheckedString == "Y")
                {
                    if (isOpen == true)
                    {
                        pGrid_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIDX_CheckBox);

                        mPrintPage++;

                        vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_WAGE_TYPE);
                        vWAGE_TYPE = ConvertString(vObject);
                        vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_PAY_TYPE);
                        vPAY_TYPE = ConvertString(vObject);

                        mCurrentRow = (mCopy_EndRow * mPageNumber) + 1;  //���� �μ� row ��ġ ���� : �μ�row�� �������� + 1 �� ������ ������ ���.

                        mCopyLineSUM = CopyAndPaste(mPrinting, mCurrentRow, mSourceSheet1, pSeal_Stamp);

                        //vRow++;
                        // �������� �μ�
                        XLLine_HEADER_SIV(pGrid_MONTH_PAYMENT, vRow, mCurrentRow, vGridCol); //Heaer �� ��������, �ѱݾ�

                        //����ٹ� �μ�
                        XLLine_OT_SIV(pIDA_MONTH_OT, mCurrentRow + 20);

                        //���°� �μ�
                        XLLine_DUTY_SIV(pIDA_MONTH_DUTY, mCurrentRow + 20);

                        //�����׸� �μ�
                        XLLine_ALLOWANCE_SIV(pIDA_PAY_ALLOWANCE, mCurrentRow + 32);

                        //�����׸� �μ�
                        XLLine_DEDUCTION_SIV(pIDA_PAY_DEDUCTION, mCurrentRow + 32);

                        // vCheckedString2 = pGrid_MONTH_PAYMENT.GetCellValue(vRow + 1, vIDX_CheckBox).ToString();


                        ////���忡 ���� �μ�////
                        ////VCC�� A5�� �μ��ϱ� ������ �ּ�//
                        //if (vCheckedString2 == "Y")
                        //{
                        //    for (int vRow2 = 0; vRow2 < 1; vRow2++)
                        //    {
                        //        vRow++;
                        //        vRowCount++;

                        //        pGrid_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIDX_CheckBox);
                        //        pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                        //        mCurrentCol = mCopy_EndCol + 1;  //���� �μ� row ��ġ ���� : �μ�row�� �������� + 1 �� ������ ������ ���.

                        //        mCopyLineSUM = CopyAndPaste2(mPrinting, mCurrentRow, mCurrentCol, mSourceSheet1);


                        //        // �������� �μ�
                        //        XLLine_HEADER2(pGrid_MONTH_PAYMENT, vRow, mCurrentRow, vGridCol); //Heaer �� ��������, �ѱݾ�

                        //        //����ٹ� �μ�
                        //        XLLine_OT2(pIDA_MONTH_OT, mCurrentRow + 16);

                        //        //���°� �μ�
                        //        XLLine_DUTY2(pIDA_MONTH_DUTY, mCurrentRow + 16);

                        //        //�����׸� �μ�
                        //        XLLine_ALLOWANCE2(pIDA_PAY_ALLOWANCE, mCurrentRow + 37);

                        //        //�����׸� �μ�
                        //        XLLine_DEDUCTION2(pIDA_PAY_DEDUCTION, mCurrentRow + 37);

                        //        //if (vTotalRow == vRowCount)
                        //        //{
                        //        //    if (pCourse == "PRINT")
                        //        //    {
                        //        //        Printing(1, mPageNumber);
                        //        //    }
                        //        //    else if (pCourse == "FILE")
                        //        //    {
                        //        //        SAVE("PAY_");
                        //        //    }
                        //        //}
                        //        //else if (mMaxPrintPage < mPageNumber)
                        //        //{
                        //        //    if (pCourse == "PRINT")
                        //        //    {
                        //        //        Printing(1, mPageNumber);

                        //        //        mPrinting.XLOpenFileClose();
                        //        //        isOpen = XLFileOpen();
                        //        //    }
                        //        //    else if (pCourse == "FILE")
                        //        //    {
                        //        //        SAVE("PAY_");

                        //        //        mPrinting.XLOpenFileClose();
                        //        //        isOpen = XLFileOpen();
                        //        //    }
                        //        //    mPageNumber = 0;
                        //        //    mCurrentRow = 1;
                        //        //    mCopyLineSUM = 1;
                        //        //}





                        //    }

                        //}
                        if (vTotalRow == vRowCount)
                        {
                            if (pOUTPUT_TYPE == "PRINT")
                            {
                                Printing(1, mPageNumber);
                            }
                            else if (pOUTPUT_TYPE == "FILE")
                            {
                                SAVE("PAY_");
                            }
                        }
                        else if (mMaxPrintPage < mPageNumber)
                        {
                            if (pOUTPUT_TYPE == "PRINT")
                            {
                                Printing(1, mPageNumber);

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();
                            }
                            else if (pOUTPUT_TYPE == "FILE")
                            {
                                SAVE("PAY_");

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();
                            }
                            mPageNumber = 0;
                            mCurrentRow = 1;
                            mCopyLineSUM = 1;
                        }




                    }

                    pGrid_MONTH_PAYMENT.SetCellValue(vRow, vIDX_CheckBox, "N");
                }
                else if (vTotalRow == vRowCount)
                {
                    if (isOpen == true)
                    {
                        if (pOUTPUT_TYPE == "PRINT")
                        {
                            Printing(1, mPageNumber);
                        }
                        else if (pOUTPUT_TYPE == "FILE")
                        {
                            SAVE("PAY_");
                        }
                    }
                }              

                
            }
            System.Windows.Forms.Application.UseWaitCursor = false;
            pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();

            return mPageNumber;
        }



        public int WriteMain_BHK(string pOUTPUT_TYPE, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_MONTH_PAYMENT
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PAY_ALLOWANCE
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_PAY_DEDUCTION
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_BONUS_ALLOWANCE
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_BONUS_DEDUCTION
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_MONTH_DUTY
                                                   , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_MONTH_OT
                                                   , string pSeal_Stamp
            )
        {
            string vMessageText = string.Empty;
            object vObject = null;
            string vWAGE_TYPE = string.Empty;
            string vPAY_TYPE = string.Empty;
            string vCheckedString = "N";
            string vCheckedString2 = "N";

            int vIDX_CheckBox = pGrid_MONTH_PAYMENT.GetColumnToIndex("SELECT_CHECK_YN");
            int vIDX_WAGE_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("WAGE_TYPE");
            int vIDX_PAY_TYPE = pGrid_MONTH_PAYMENT.GetColumnToIndex("PAY_TYPE");

            int[] vGridCol;

            bool isOpen = XLFileOpen();

            mPageNumber = 0;
            mCopyLineSUM = 1;

            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 48;
            mCopy_EndRow = 76;

            mCurrentRow = 1;
            mCurrentCol = 1;

            int[] vGDColumn_1;
            int[] vXLColumn_1;

            int[] vGDColumn_2;
            int[] vXLColumn_2;

            int[] vGDColumn_3;
            int[] vXLColumn_3;

            int[] vGDColumn_4;
            int[] vXLColumn_4;

            int[] vGDColumn_5;
            int[] vXLColumn_5;

            int[] vGDColumn_6;
            int[] vXLColumn_6;

            int vTotalRow = pGrid_MONTH_PAYMENT.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 0;

            int vSecondPrinting = 29; //30��°�� �μ�
            int vCountPrinting = 0;

            SetArray1(pIDA_PAY_ALLOWANCE.OraSelectData, out vGDColumn_1, out vXLColumn_1);
            SetArray2(pIDA_PAY_DEDUCTION.OraSelectData, out vGDColumn_2, out vXLColumn_2);
            SetArray3(pIDA_BONUS_ALLOWANCE.OraSelectData, out vGDColumn_3, out vXLColumn_3);
            SetArray4(pIDA_BONUS_DEDUCTION.OraSelectData, out vGDColumn_4, out vXLColumn_4);
            SetArray5_BHK(pIDA_MONTH_DUTY.OraSelectData, out vGDColumn_5, out vXLColumn_5);
            SetArray6_BHK(pGrid_MONTH_PAYMENT, out vGDColumn_6, out vXLColumn_6);

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;

                vMessageText = string.Format("Printing Rows : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);

                System.Windows.Forms.Application.UseWaitCursor = true;
                pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                System.Windows.Forms.Application.DoEvents();

                vCheckedString = Convert.ToString(pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_CheckBox));
                if (vCheckedString == "Y")
                {
                    if (isOpen == true)
                    {
                        pGrid_MONTH_PAYMENT.CurrentCellMoveTo(vRow, vIDX_CheckBox);

                        mPrintPage++;

                        vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_WAGE_TYPE);
                        vWAGE_TYPE = ConvertString(vObject);
                        vObject = pGrid_MONTH_PAYMENT.GetCellValue(vRow, vIDX_PAY_TYPE);
                        vPAY_TYPE = ConvertString(vObject);

                        mCurrentRow = (mCopy_EndRow * mPageNumber) + 1;  //���� �μ� row ��ġ ���� : �μ�row�� �������� + 1 �� ������ ������ ���.

                        if (vWAGE_TYPE == "P1" && (vPAY_TYPE == "2" || vPAY_TYPE == "4"))
                        {
                            mCopyLineSUM = CopyAndPaste_BHK(mPrinting, mCopyLineSUM, "DAILY", pSeal_Stamp);
                            vPrintingLine = (mCopyLineSUM - mCopy_EndRow) + (mPrintingLineSTART - 1);

                            //������
                            int vLinePrinting_1 = vPrintingLine + 28;
                            for (int vRow1 = 0; vRow1 < pIDA_PAY_ALLOWANCE.OraSelectData.Rows.Count; vRow1++)
                            {
                                vLinePrinting_1 = XLLine_1(pIDA_PAY_ALLOWANCE.OraSelectData.Rows[vRow1], vLinePrinting_1, vGDColumn_1, vXLColumn_1); //�޿� ����
                            }

                            int vLinePrinting_2 = vPrintingLine + 28;
                            for (int vRow2 = 0; vRow2 < pIDA_PAY_DEDUCTION.OraSelectData.Rows.Count; vRow2++)
                            {
                                vLinePrinting_2 = XLLine_2(pIDA_PAY_DEDUCTION.OraSelectData.Rows[vRow2], vLinePrinting_2, vGDColumn_2, vXLColumn_2); //�޿� ����
                            }

                            int vLinePrinting_3 = vPrintingLine + 56;
                            for (int vRow3 = 0; vRow3 < pIDA_BONUS_ALLOWANCE.OraSelectData.Rows.Count; vRow3++)
                            {
                                vLinePrinting_3 = XLLine_3(pIDA_BONUS_ALLOWANCE.OraSelectData.Rows[vRow3], vLinePrinting_3, vGDColumn_3, vXLColumn_3); //�� ����
                            }

                            int vLinePrinting_4 = vPrintingLine + 56;
                            for (int vRow4 = 0; vRow4 < pIDA_BONUS_DEDUCTION.OraSelectData.Rows.Count; vRow4++)
                            {
                                vLinePrinting_4 = XLLine_4(pIDA_BONUS_DEDUCTION.OraSelectData.Rows[vRow4], vLinePrinting_4, vGDColumn_4, vXLColumn_4); //�� ����
                            }

                            int vLinePrinting_5 = vPrintingLine + 20;
                            for (int vRow5 = 0; vRow5 < pIDA_MONTH_DUTY.OraSelectData.Rows.Count; vRow5++)
                            {
                                vLinePrinting_5 = XLLine_5_BHK(pIDA_MONTH_DUTY.OraSelectData.Rows[vRow5], vLinePrinting_5, vGDColumn_5, vXLColumn_5); //�ٹ��ð� �� �ΰ�����
                            }

                            vPrintingLine = XLLine_6_BHK(pGrid_MONTH_PAYMENT, vRow, vPrintingLine, vGDColumn_6, vXLColumn_6, "DAILY"); //Heaer �� ��������, �ѱݾ�
                        }
                        else
                        {
                            mCopyLineSUM = CopyAndPaste_BHK(mPrinting, mCopyLineSUM, "MONTH", pSeal_Stamp);
                            vPrintingLine = (mCopyLineSUM - mCopy_EndRow) + (mPrintingLineSTART - 1);

                            //������
                            int vLinePrinting_1 = vPrintingLine + 28;
                            for (int vRow1 = 0; vRow1 < pIDA_PAY_ALLOWANCE.OraSelectData.Rows.Count; vRow1++)
                            {
                                vLinePrinting_1 = XLLine_1(pIDA_PAY_ALLOWANCE.OraSelectData.Rows[vRow1], vLinePrinting_1, vGDColumn_1, vXLColumn_1); //�޿� ����
                            }

                            int vLinePrinting_2 = vPrintingLine + 28;
                            for (int vRow2 = 0; vRow2 < pIDA_PAY_DEDUCTION.OraSelectData.Rows.Count; vRow2++)
                            {
                                vLinePrinting_2 = XLLine_2(pIDA_PAY_DEDUCTION.OraSelectData.Rows[vRow2], vLinePrinting_2, vGDColumn_2, vXLColumn_2); //�޿� ����
                            }

                            int vLinePrinting_5 = vPrintingLine + 20;
                            for (int vRow5 = 0; vRow5 < pIDA_MONTH_DUTY.OraSelectData.Rows.Count; vRow5++)
                            {
                                vLinePrinting_5 = XLLine_5_BHK(pIDA_MONTH_DUTY.OraSelectData.Rows[vRow5], vLinePrinting_5, vGDColumn_5, vXLColumn_5); //�ٹ��ð� �� �ΰ�����
                            }

                            vPrintingLine = XLLine_6_BHK(pGrid_MONTH_PAYMENT, vRow, vPrintingLine, vGDColumn_6, vXLColumn_6, "MONTH"); //Heaer �� ��������, �ѱݾ�
                        }

                        if (vTotalRow == vRowCount)
                        {
                            if (pOUTPUT_TYPE == "PRINT")
                            {
                                Printing(1, mPageNumber);
                            }
                            else if (pOUTPUT_TYPE == "PDF")
                            {
                                PDF_Save("PAY_");
                            }
                            else if (pOUTPUT_TYPE == "FILE")
                            {
                                SAVE("PAY_");
                            }
                        }
                        else if (mMaxPrintPage < mPageNumber)
                        {
                            if (pOUTPUT_TYPE == "PRINT")
                            {
                                Printing(1, mPageNumber);

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();
                            }
                            else if (pOUTPUT_TYPE == "PDF")
                            {
                                PDF_Save("PAY_");
                            }
                            else if (pOUTPUT_TYPE == "FILE")
                            {
                                SAVE("PAY_");

                                mPrinting.XLOpenFileClose();
                                isOpen = XLFileOpen();
                            }
                            mPageNumber = 0;
                            mCurrentRow = 1;
                            mCopyLineSUM = 1;
                        }
                    }
                    pGrid_MONTH_PAYMENT.SetCellValue(vRow, vIDX_CheckBox, "N");
                }
                else if (vTotalRow == vRowCount)
                {
                    if (isOpen == true)
                    {
                        if (pOUTPUT_TYPE == "PRINT")
                        {
                            Printing(1, mPageNumber);
                        }
                        else if (pOUTPUT_TYPE == "FILE")
                        {
                            SAVE("PAY_");
                        }
                    }
                }
            }

            System.Windows.Forms.Application.UseWaitCursor = false;
            pGrid_MONTH_PAYMENT.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();

            return mPageNumber;
        }

        #endregion;

        #endregion;

        #region ----- Copy&Paste Sheet Method ---- 

        //ù��° ������ ����
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCurrentRow, string pSourceSheet, string pSeal_Stamp)
        {
            //int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            if(pSeal_Stamp == "Y")
            {
                mPrinting.XLDeleteBarCode(pIndexImage: 2);
            }

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(pSourceSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pCurrentRow, mCopy_StartCol, pCurrentRow + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber++; //������ ��ȣ

            return pCurrentRow + mCopy_EndRow;
        }

        private int CopyAndPaste2(XL.XLPrint pPrinting, int pCurrentRow, int pCurrentCol, string pSourceSheet)
        {
            //int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(pSourceSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pCurrentRow, pCurrentCol, pCurrentRow + mCopy_EndRow, pCurrentCol + mCopy_EndCol + 1);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);



            return pCurrentCol + mCopy_EndCol;
        }

        //ù��° ������ ����
        private int CopyAndPaste_BHK(XL.XLPrint pPrinting, int pCurrentRow, string pCourse, string vSeal_Flag)
        {
            int vCopySumPrintingLine = pCurrentRow;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;

            if (pCourse == "DAILY")
            {
                pPrinting.XLActiveSheet(mSourceSheet2);
                if (vSeal_Flag == "N")
                {
                    mPrinting.XLDeleteBarCode(pIndexImage: 3);
                }
            }
            else if (pCourse == "MONTH")
            {
                pPrinting.XLActiveSheet(mSourceSheet1);
                if (vSeal_Flag == "N")
                {
                    mPrinting.XLDeleteBarCode(pIndexImage: 3);
                }
            }
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pCurrentRow, 1, pCurrentRow + mCopy_EndRow, mCopy_EndCol + 1);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            int vCopy_EndRow = pCurrentRow + mCopy_EndRow;
            mPrinting.XLHPageBreaks_Add(mPrinting.XLGetRange("A" + vCopy_EndRow));

            mPageNumber++; //������ ��ȣ

            return pCurrentRow + mCopy_EndRow;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLDeleteSheet("Source1");
            mPrinting.XLPrinting(pPageSTART, pPageEND);
            
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xlsx", vWallpaperFolder.ToString(), vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        public void PDF_Save(string pSaveFileName)
        {
            if (pSaveFileName == string.Empty)
            {
                return;
            }
            mPrinting.XLDeleteSheet(mSourceSheet1);
            mPrinting.XLDeleteSheet(mSourceSheet2);

            ////��ȣ�� �ּ�
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D2}.pdf", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.pdf", vWallpaperFolder, vSaveFileName);

            mPrinting.XLSaveAs_PDF(pSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
        }

        public bool SaveFile(string pSaveFileName, string pPassword)
        {
            try
            {
                mPrinting.XLDeleteSheet(mSourceSheet1);
            }
            catch (Exception Ex)
            {
                mAppInterface.OnAppMessageEvent("DeleteSheet Error : " + Ex.Message);
                return false;
            }
            try
            {
                mPrinting.XLDeleteSheet(mSourceSheet2);
            }
            catch (Exception Ex)
            {
                mAppInterface.OnAppMessageEvent("DeleteSheet Error : " + Ex.Message);
                return false;
            }
            try
            {
                mPrinting.XLSaveAs_PDF(pSaveFileName, pPassword);
            }
            catch (Exception Ex)
            {
                mAppInterface.OnAppMessageEvent("SaveFile Error : " + Ex.Message);
                return false;
            }
            return true;
        }

        #endregion;
    }
}