using System;

namespace HRMF0537
{
    public class XLPrinting
    {
        #region ----- Variables -----

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv mAppInterfaceAdv = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mCopySumPrintingLine = 1; //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ
        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;
        private int mPrintingLineMAX = 43; //43:12����� 43�����, 32:�ݺ��Ǵ� ����
        private int mIncrementCopyMAX = 45;
        private int mPositionPrintLineSTART = 12; //���� ��½� ���� ���� �� ��ġ ����

        private string mCorporationName = string.Empty;
        private string mUserName = string.Empty;
        private string mYYYYMM = string.Empty;
        private string mWageTypeName = string.Empty;
        private string mDepartmentName = string.Empty;
        private string mPringingDateTime = string.Empty;

        private string mPageString = string.Empty;
        private int mPageTotalNumber = 0;
        private int mCountPage = 0;

        private string[] mGridColumn;
        private int[] mXLColumn;
        private decimal[] mSumValueColumn;

        private string mDepartment = string.Empty;

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

        public int PrintingLineMAX
        {
            set
            {
                mPrintingLineMAX = value;
            }
        }

        public int IncrementCopyMAX
        {
            set
            {
                mIncrementCopyMAX = value;
            }
        }

        public int PositionPrintLineSTART
        {
            set
            {
                mPositionPrintLineSTART = value;
            }
        }

        public int CopySumPrintingLine
        {
            set
            {
                mCopySumPrintingLine = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterfaceAdv pAppInterfaceAdv)
        {
            mPrinting = new XL.XLPrint();
            mAppInterfaceAdv = pAppInterfaceAdv;
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

                int vCutRight = 2;
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

        #region ----- Line Clear All Methods ----

        private void XlAllLineClear(int[] pXLColumn)
        {
            object vObject = null;
            int vPrintingLineMAX = mPrintingLineMAX + 1;
            int mXLColumnCount = pXLColumn.Length;

            mPrinting.XLActiveSheet("SourceTab1");

            for (int vXLine = mPositionPrintLineSTART; vXLine < vPrintingLineMAX; vXLine++)
            {
                for (int vCOL = 1; vCOL < mXLColumnCount; vCOL++)
                {
                    mPrinting.XLSetCell(vXLine, pXLColumn[vCOL], vObject);
                }
            }
        }

        #endregion;

        #region ----- Excel Wirte [Header] Methods ----

        public void HeaderWrite(string pUserName, string pPrintingDateTime, string pYYYYMM, string pWageTypeName, string pDepartment_NAME, string pPageString, string pCorporationName)
        {
            bool isNull = false;
            try
            {
                System.Drawing.Point vCellPoint01 = new System.Drawing.Point(2, 2);    //Title
                System.Drawing.Point vCellPoint02 = new System.Drawing.Point(4, 6);    //�����
                System.Drawing.Point vCellPoint03 = new System.Drawing.Point(5, 6);    //�޿�����
                System.Drawing.Point vCellPoint04 = new System.Drawing.Point(5, 19);   //�μ�
                System.Drawing.Point vCellPoint05 = new System.Drawing.Point(4, 56);   //������
                System.Drawing.Point vCellPoint06 = new System.Drawing.Point(5, 56);   //�������
                System.Drawing.Point vCellPoint07 = new System.Drawing.Point(44, 41);  //��ü

                mPrinting.XLActiveSheet("SourceTab1"); //���� ���ڸ� �ֱ� ���� ��Ʈ ����

                //Title
                isNull = string.IsNullOrEmpty(pYYYYMM);
                if (isNull != true)
                {
                    string vYear = pYYYYMM.Substring(0, 4);
                    string vMonth = pYYYYMM.Substring(5, 2);
                    string vTitle = string.Format("{0}�� {1}�� �޿� ����", vYear, vMonth);
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, vTitle);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint01.X, vCellPoint01.Y, null);
                }

                //�����
                isNull = string.IsNullOrEmpty(pUserName);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, pUserName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint02.X, vCellPoint02.Y, null);
                }

                //�޿�����
                isNull = string.IsNullOrEmpty(pWageTypeName);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, pWageTypeName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint03.X, vCellPoint03.Y, "��ü");
                }

                //�μ�
                isNull = string.IsNullOrEmpty(pDepartment_NAME);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, pDepartment_NAME);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint04.X, vCellPoint04.Y, "��ü");
                }

                //������
                isNull = string.IsNullOrEmpty(pPageString);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, pPageString);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint05.X, vCellPoint05.Y, null);
                }

                //�������
                isNull = string.IsNullOrEmpty(pPrintingDateTime);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, pPrintingDateTime);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint06.X, vCellPoint06.Y, null);
                }

                //��ü
                isNull = string.IsNullOrEmpty(pCorporationName);
                if (isNull != true)
                {
                    mPrinting.XLSetCell(vCellPoint07.X, vCellPoint07.Y, pCorporationName);
                }
                else
                {
                    mPrinting.XLSetCell(vCellPoint07.X, vCellPoint07.Y, null);
                }
            }
            catch (System.Exception ex)
            {
                mAppInterfaceAdv.OnAppMessage(ex.Message);

                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }
        }

        #endregion;

        #region ----- Line SLIP Methods ----

        #region ----- Array Set ----

        private void SetArray(out string[] pGridColumn, out int[] pXLColumn)
        {
            mSumValueColumn = new decimal[69];
            pGridColumn = new string[69];
            pXLColumn = new int[69];

            pGridColumn[01] = "DEPT_NAME";                    //�μ�
            pGridColumn[02] = "PERSON_NUM";                   //�����ȣ
            pGridColumn[03] = ""; //�⺻��
            pGridColumn[04] = "TOTAL_ATT_DAY"; //����ٹ�
            pGridColumn[05] = "HOLY_1_TIME"; //����Ư��
            pGridColumn[06] = "DUTY_30"; //�ٹ�(����)
            pGridColumn[07] = "TOT_DED_COUNT"; //�̱ٹ�
            pGridColumn[08] = "A01";                          //�⺻��[�����׸�]
            pGridColumn[09] = "A08";                          //��������
            pGridColumn[10] = "A02";                          //��å����
            pGridColumn[11] = "A03";                          //�ټӼ���
            pGridColumn[12] = "A18"; //��������
            pGridColumn[13] = "D01";                          //�ҵ漼
            pGridColumn[14] = "D02";                          //�ֹμ�
            pGridColumn[15] = "D03";                          //���ο���
            pGridColumn[16] = "D05";                          //�ǰ�����
            pGridColumn[17] = "TOT_SUPPLY_AMOUNT";            //�����޾�[�����޾�]
            pGridColumn[18] = "POST_NAME";                    //����
            pGridColumn[19] = "NAME";                         //����
            pGridColumn[20] = ""; //�ϱ�
            pGridColumn[21] = "OVER_TIME"; //����ٷ�(����ð�)
            pGridColumn[22] = "HOLY_1_OT"; //���Ͽ���
            pGridColumn[23] = "S_HOLY_1_COUNT"; //����
            pGridColumn[24] = "WEEKLY_DED_COUNT"; //������
            pGridColumn[25] = "A06";                          //�ڰݼ���
            pGridColumn[26] = "A11";                          //�ð��ܼ���
            pGridColumn[27] = "A12";                          //�������
            pGridColumn[28] = "A13";                          //�߰�����
            pGridColumn[29] = "A14";                          //Ư�ټ���
            pGridColumn[30] = "D04";                          //��뺸��
            pGridColumn[31] = "";                          //���ȸ
            pGridColumn[32] = "D06"; //�Ĵ�
            pGridColumn[33] = ""; //���ұ�
            pGridColumn[34] = "TOT_DED_AMOUNT";               //�Ѱ�����[�Ѱ�����]
            pGridColumn[35] = ""; //�����Ի���
            pGridColumn[36] = "WAGE_TYPE_NAME";               //�޻󿩱���
            pGridColumn[37] = ""; //�ñ�
            pGridColumn[38] = "NIGHT_BONUS_TIME"; //�߰��ٷ�(�߰��ð�)
            pGridColumn[39] = "HOLY_1_NIGHT"; //���Ͼ߰�
            pGridColumn[40] = "HOLY_1_COUNT"; //����
            pGridColumn[41] = "DUTY_30"; //����
            pGridColumn[42] = "";                          //���ټ���
            pGridColumn[43] = "";                          //������������
            pGridColumn[44] = "";                          //���
            pGridColumn[45] = "A25";                          //����������
            pGridColumn[46] = "A07";                          //��Ÿ����
            pGridColumn[47] = ""; //�������
            pGridColumn[48] = ""; //ī���
            pGridColumn[49] = "D14";                          //��Ÿ
            pGridColumn[50] = ""; //��ȣȸ
            pGridColumn[51] = "REAL_AMOUNT";                  //�����޾�[�����޾�]
            pGridColumn[52] = ""; //
            pGridColumn[53] = ""; //
            pGridColumn[54] = ""; //
            pGridColumn[55] = ""; //
            pGridColumn[56] = "LATE_TIME"; //���°���
            pGridColumn[57] = "HOLY_0_COUNT"; //����
            pGridColumn[58] = ""; //���ñ�
            pGridColumn[59] = ""; //
            pGridColumn[60] = "A09";                          //�󿩱�
            pGridColumn[61] = "";                          //�߰��������
            pGridColumn[62] = "A10";                          //�����ұ޺�
            pGridColumn[63] = ""; //�����ް�
            pGridColumn[64] = "D16";                          //����ҵ漼
            pGridColumn[65] = "D17";                          //�����ֹμ�
            pGridColumn[66] = ""; //���ο��ݼұ޺�
            pGridColumn[67] = ""; //�ǰ���������
            pGridColumn[68] = ""; //������Ѿ�

            pXLColumn[01] = 2;  //�μ�
            pXLColumn[02] = 5;  //�����ȣ
            pXLColumn[03] = 8;  //�⺻��
            pXLColumn[04] = 11; //����ٹ�
            pXLColumn[05] = 14; //����Ư��
            pXLColumn[06] = 17; //�ٹ�(����)
            pXLColumn[07] = 20; //�̱ٹ�
            pXLColumn[08] = 23; //�⺻��[�����׸�]
            pXLColumn[09] = 27; //��������
            pXLColumn[10] = 31; //��å����
            pXLColumn[11] = 35; //�ټӼ���
            pXLColumn[12] = 39; //��������
            pXLColumn[13] = 43; //�ҵ漼
            pXLColumn[14] = 47; //�ֹμ�
            pXLColumn[15] = 51; //���ο���
            pXLColumn[16] = 55; //�ǰ�����
            pXLColumn[17] = 59; //�����޾�[�����޾�]
            pXLColumn[18] = 2;  //����
            pXLColumn[19] = 5;  //����
            pXLColumn[20] = 8;  //�ϱ�
            pXLColumn[21] = 11; //����ٷ�
            pXLColumn[22] = 14; //���Ͽ���
            pXLColumn[23] = 17; //����
            pXLColumn[24] = 20; //������
            pXLColumn[25] = 23; //�ڰݼ���
            pXLColumn[26] = 27; //�ð��ܼ���
            pXLColumn[27] = 31; //��������
            pXLColumn[28] = 35; //�߰�����
            pXLColumn[29] = 39; //Ư�ټ���
            pXLColumn[30] = 43; //��뺸��
            pXLColumn[31] = 47; //���ȸ
            pXLColumn[32] = 51; //�Ĵ�
            pXLColumn[33] = 55; //���ұ�
            pXLColumn[34] = 59; //�Ѱ�����[�Ѱ�����]
            pXLColumn[35] = 2;  //�����Ի���
            pXLColumn[36] = 5;  //�޻󿩱���
            pXLColumn[37] = 8;  //�ñ�
            pXLColumn[38] = 11; //�߰��ٷ�
            pXLColumn[39] = 14; //���Ͼ߰�
            pXLColumn[40] = 17; //����
            pXLColumn[41] = 20; //����
            pXLColumn[42] = 23; //���ټ���
            pXLColumn[43] = 27; //������������
            pXLColumn[44] = 31; //���
            pXLColumn[45] = 35; //����������
            pXLColumn[46] = 39; //��Ÿ����
            pXLColumn[47] = 43; //�������
            pXLColumn[48] = 47; //ī���
            pXLColumn[49] = 51; //��Ÿ
            pXLColumn[50] = 55; //��ȣȸ
            pXLColumn[51] = 59; //�����޾�[�����޾�]
            pXLColumn[52] = 2;  //
            pXLColumn[53] = 5;  //
            pXLColumn[54] = 8;  //
            pXLColumn[55] = 11; //
            pXLColumn[56] = 14; //���°���
            pXLColumn[57] = 17; //����
            pXLColumn[58] = 20; //���ñ�
            pXLColumn[59] = 23; //
            pXLColumn[60] = 27; //�󿩱�
            pXLColumn[61] = 31; //�߰��������
            pXLColumn[62] = 35; //�����ұ޺�
            pXLColumn[63] = 39; //�����ް�
            pXLColumn[64] = 43; //����ҵ漼
            pXLColumn[65] = 47; //�����ֹμ�
            pXLColumn[66] = 51; //���ο��ݼұ޺�
            pXLColumn[67] = 55; //�ǰ���������
            pXLColumn[68] = 59; //������Ѿ�
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

        #region ----- Convert DateTime Methods ----

        private string ConvertDateTime(object pObject)
        {
            string vTextDateTimeLong = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd HH:mm:ss", null);
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vTextDateTimeLong;
        }

        private string ConvertDate(object pObject)
        {
            string vTextDateTimeShort = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        vTextDateTimeShort = vDateTime.ToShortDateString();
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vTextDateTimeShort;
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
                mAppInterfaceAdv.OnAppMessage(mMessageError);
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
                mAppInterfaceAdv.OnAppMessage(mMessageError);
            }

            return vIsConvert;
        }

        private bool IsConvertNumber(string pStringNumber, out decimal pConvertDecimal)
        {
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pStringNumber != null)
                {
                    decimal vIsConvertNum = decimal.Parse(pStringNumber);
                    pConvertDecimal = vIsConvertNum;
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterfaceAdv.OnAppMessage(mMessageError);
            }

            return vIsConvert;
        }

        private bool IsConvertDate(object pObject, out string pConvertDateTimeShort)
        {
            bool vIsConvert = false;
            pConvertDateTimeShort = string.Empty;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime.ToShortDateString();
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vIsConvert;
        }

        #endregion;

        #region ----- Xl Clear Value SUM Methods -----

        private void ClearValueSumValue()
        {
            int vCountIndex = mSumValueColumn.Length;

            for (int vRow = 0; vRow < vCountIndex; vRow++)
            {
                mSumValueColumn[vRow] = 0m;
            }
        }

        #endregion;

        #region ----- Xl SUM Methods -----

        private int XLSUM(int pPrintingLine, int[] pXLColumn, string pDepartment, decimal[] pSumValueColumn)
        {
            int vXLine = pPrintingLine; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLIndexColumn = 0;
            string vConvertString = string.Empty;
            decimal vSumValue = 0m;

            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                //[01]
                vXLIndexColumn = pXLColumn[1];
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, pDepartment);

                //[03]
                vXLIndexColumn = pXLColumn[3];
                vSumValue = pSumValueColumn[3];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[04]
                vXLIndexColumn = pXLColumn[4];
                vSumValue = pSumValueColumn[4];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[05]
                vXLIndexColumn = pXLColumn[5];
                vSumValue = pSumValueColumn[5];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[06]
                vXLIndexColumn = pXLColumn[6];
                vSumValue = pSumValueColumn[6];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[07]
                vXLIndexColumn = pXLColumn[7];
                vSumValue = pSumValueColumn[7];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[08]
                vXLIndexColumn = pXLColumn[8];
                vSumValue = pSumValueColumn[8];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[09]
                vXLIndexColumn = pXLColumn[9];
                vSumValue = pSumValueColumn[9];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[10]
                vXLIndexColumn = pXLColumn[10];
                vSumValue = pSumValueColumn[10];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[11]
                vXLIndexColumn = pXLColumn[11];
                vSumValue = pSumValueColumn[11];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[12]
                vXLIndexColumn = pXLColumn[12];
                vSumValue = pSumValueColumn[12];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[13]
                vXLIndexColumn = pXLColumn[13];
                vSumValue = pSumValueColumn[13];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[14]
                vXLIndexColumn = pXLColumn[14];
                vSumValue = pSumValueColumn[14];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[15]
                vXLIndexColumn = pXLColumn[15];
                vSumValue = pSumValueColumn[15];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[16]
                vXLIndexColumn = pXLColumn[16];
                vSumValue = pSumValueColumn[16];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[17]
                vXLIndexColumn = pXLColumn[17];
                vSumValue = pSumValueColumn[17];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[20]
                vXLIndexColumn = pXLColumn[20];
                vSumValue = pSumValueColumn[20];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[21]
                vXLIndexColumn = pXLColumn[21];
                vSumValue = pSumValueColumn[21];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[22]
                vXLIndexColumn = pXLColumn[22];
                vSumValue = pSumValueColumn[22];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[23]
                vXLIndexColumn = pXLColumn[23];
                vSumValue = pSumValueColumn[23];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[24]
                vXLIndexColumn = pXLColumn[24];
                vSumValue = pSumValueColumn[24];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[25]
                vXLIndexColumn = pXLColumn[25];
                vSumValue = pSumValueColumn[25];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[26]
                vXLIndexColumn = pXLColumn[26];
                vSumValue = pSumValueColumn[26];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[27]
                vXLIndexColumn = pXLColumn[27];
                vSumValue = pSumValueColumn[27];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[28]
                vXLIndexColumn = pXLColumn[28];
                vSumValue = pSumValueColumn[28];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[29]
                vXLIndexColumn = pXLColumn[29];
                vSumValue = pSumValueColumn[29];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[30]
                vXLIndexColumn = pXLColumn[30];
                vSumValue = pSumValueColumn[30];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[31]
                vXLIndexColumn = pXLColumn[31];
                vSumValue = pSumValueColumn[31];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[32]
                vXLIndexColumn = pXLColumn[32];
                vSumValue = pSumValueColumn[32];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[33]
                vXLIndexColumn = pXLColumn[33];
                vSumValue = pSumValueColumn[33];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[34]
                vXLIndexColumn = pXLColumn[34];
                vSumValue = pSumValueColumn[34];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[37]
                vXLIndexColumn = pXLColumn[37];
                vSumValue = pSumValueColumn[37];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[38]
                vXLIndexColumn = pXLColumn[38];
                vSumValue = pSumValueColumn[38];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[39]
                vXLIndexColumn = pXLColumn[39];
                vSumValue = pSumValueColumn[39];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[40]
                vXLIndexColumn = pXLColumn[40];
                vSumValue = pSumValueColumn[40];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[41]
                vXLIndexColumn = pXLColumn[41];
                vSumValue = pSumValueColumn[41];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[42]
                vXLIndexColumn = pXLColumn[42];
                vSumValue = pSumValueColumn[42];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[43]
                vXLIndexColumn = pXLColumn[43];
                vSumValue = pSumValueColumn[43];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[44]
                vXLIndexColumn = pXLColumn[44];
                vSumValue = pSumValueColumn[44];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[45]
                vXLIndexColumn = pXLColumn[45];
                vSumValue = pSumValueColumn[45];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[46]
                vXLIndexColumn = pXLColumn[46];
                vSumValue = pSumValueColumn[46];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[47]
                vXLIndexColumn = pXLColumn[47];
                vSumValue = pSumValueColumn[47];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[48]
                vXLIndexColumn = pXLColumn[48];
                vSumValue = pSumValueColumn[48];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[49]
                vXLIndexColumn = pXLColumn[49];
                vSumValue = pSumValueColumn[49];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[50]
                vXLIndexColumn = pXLColumn[50];
                vSumValue = pSumValueColumn[50];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[51]
                vXLIndexColumn = pXLColumn[51];
                vSumValue = pSumValueColumn[51];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[54]
                vXLIndexColumn = pXLColumn[54];
                vSumValue = pSumValueColumn[54];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[55]
                vXLIndexColumn = pXLColumn[55];
                vSumValue = pSumValueColumn[55];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[56]
                vXLIndexColumn = pXLColumn[56];
                vSumValue = pSumValueColumn[56];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[57]
                vXLIndexColumn = pXLColumn[57];
                vSumValue = pSumValueColumn[57];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[58]
                vXLIndexColumn = pXLColumn[58];
                vSumValue = pSumValueColumn[58];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[59]
                vXLIndexColumn = pXLColumn[59];
                vSumValue = pSumValueColumn[59];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[60]
                vXLIndexColumn = pXLColumn[60];
                vSumValue = pSumValueColumn[60];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[61]
                vXLIndexColumn = pXLColumn[61];
                vSumValue = pSumValueColumn[61];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[62]
                vXLIndexColumn = pXLColumn[62];
                vSumValue = pSumValueColumn[62];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[63]
                vXLIndexColumn = pXLColumn[63];
                vSumValue = pSumValueColumn[63];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[64]
                vXLIndexColumn = pXLColumn[64];
                vSumValue = pSumValueColumn[64];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[65]
                vXLIndexColumn = pXLColumn[65];
                vSumValue = pSumValueColumn[65];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[66]
                vXLIndexColumn = pXLColumn[66];
                vSumValue = pSumValueColumn[66];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[67]
                vXLIndexColumn = pXLColumn[67];
                vSumValue = pSumValueColumn[67];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                //[68]
                vXLIndexColumn = pXLColumn[68];
                vSumValue = pSumValueColumn[68];
                vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vSumValue);
                mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                vXLine++;
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterfaceAdv.OnAppMessage(mMessageError);
            }


            pPrintingLine = vXLine;
            IsNewPage(pPrintingLine);
            if (mIsNewPage == true)
            {
                pPrintingLine = mPositionPrintLineSTART;
            }

            return pPrintingLine;
        }

        #endregion;

        #region ----- XlLine Methods -----

        private int XlLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pRow, int pPrintingLine, string[] pGridColumn, int[] pXLColumn)
        {
            bool vIsValueViewTemp = true;
            int vXLine = pPrintingLine; //������ ������ ǥ�õǴ� �� ��ȣ

            object vGetValue = null;
            int vGridIndexColumn = 0;
            int vXLIndexColumn = 0;

            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                //[01]
                vXLIndexColumn = pXLColumn[1];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[1]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[01]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[01]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[02]
                vXLIndexColumn = pXLColumn[2];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[2]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[02]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[02]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[03]
                vXLIndexColumn = pXLColumn[3];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[3]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[03]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[3] = mSumValueColumn[3] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[03]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[03]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[04]
                vXLIndexColumn = pXLColumn[4];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[4]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[04]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[4] = mSumValueColumn[4] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[04]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[04]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[05]
                vXLIndexColumn = pXLColumn[5];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[5]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[05]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[5] = mSumValueColumn[5] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[05]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[05]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[06]
                vXLIndexColumn = pXLColumn[6];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[6]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[06]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[6] = mSumValueColumn[6] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[06]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[06]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[07]
                vXLIndexColumn = pXLColumn[7];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[7]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[07]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[7] = mSumValueColumn[7] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[07]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[07]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[08]
                vXLIndexColumn = pXLColumn[8];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[8]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[8] = mSumValueColumn[8] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[08]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[08]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[09]
                vXLIndexColumn = pXLColumn[9];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[9]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[9] = mSumValueColumn[9] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[09]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[09]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[10]
                vXLIndexColumn = pXLColumn[10];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[10]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[10] = mSumValueColumn[10] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[10]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[10]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[11]
                vXLIndexColumn = pXLColumn[11];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[11]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[11] = mSumValueColumn[11] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[11]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[11]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[12]
                vXLIndexColumn = pXLColumn[12];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[12]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[12] = mSumValueColumn[12] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[12]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[12]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[13]
                vXLIndexColumn = pXLColumn[13];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[13]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[13] = mSumValueColumn[13] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[13]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[13]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[14]
                vXLIndexColumn = pXLColumn[14];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[14]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[14] = mSumValueColumn[14] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[14]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[14]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[15]
                vXLIndexColumn = pXLColumn[15];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[15]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[15] = mSumValueColumn[15] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[15]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[15]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[16]
                vXLIndexColumn = pXLColumn[16];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[16]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[16] = mSumValueColumn[16] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[16]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[16]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[17]
                vXLIndexColumn = pXLColumn[17];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[17]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[17] = mSumValueColumn[17] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);

                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[17]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[17]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[18]
                vXLIndexColumn = pXLColumn[18];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[18]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[18]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[18]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[19]
                vXLIndexColumn = pXLColumn[19];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[19]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[19]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[19]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[20]
                vXLIndexColumn = pXLColumn[20];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[20]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[20]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[20] = mSumValueColumn[20] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[20]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[20]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[21]
                vXLIndexColumn = pXLColumn[21];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[21]);
                ////vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[21]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[21] = mSumValueColumn[21] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[21]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[21]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[22]
                vXLIndexColumn = pXLColumn[22];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[22]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[22]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[22] = mSumValueColumn[22] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[22]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[22]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[23]
                vXLIndexColumn = pXLColumn[23];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[23]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[23]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[23] = mSumValueColumn[23] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[23]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[23]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[24]
                vXLIndexColumn = pXLColumn[24];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[24]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[24]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[24] = mSumValueColumn[24] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[24]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[24]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[25]
                vXLIndexColumn = pXLColumn[25];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[25]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[25] = mSumValueColumn[25] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[25]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[25]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[26]
                vXLIndexColumn = pXLColumn[26];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[26]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[26] = mSumValueColumn[26] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[26]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[26]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[27]
                vXLIndexColumn = pXLColumn[27];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[27]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[27] = mSumValueColumn[27] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[27]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[27]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[28]
                vXLIndexColumn = pXLColumn[28];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[28]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[28] = mSumValueColumn[28] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[28]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[28]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[29]
                vXLIndexColumn = pXLColumn[29];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[29]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[29] = mSumValueColumn[29] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[29]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[29]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[30]
                vXLIndexColumn = pXLColumn[30];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[30]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[30] = mSumValueColumn[30] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[30]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[30]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[31]
                vXLIndexColumn = pXLColumn[31];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[31]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[31] = mSumValueColumn[31] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[31]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[31]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[32]
                vXLIndexColumn = pXLColumn[32];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[32]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[32] = mSumValueColumn[32] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[32]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[32]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[33]
                vXLIndexColumn = pXLColumn[33];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[33]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[33] = mSumValueColumn[33] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[33]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[33]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[34]
                vXLIndexColumn = pXLColumn[34];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[34]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[34] = mSumValueColumn[34] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[34]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[34]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[35]
                vXLIndexColumn = pXLColumn[35];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[35]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertDate(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[35]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[35]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[36]
                vXLIndexColumn = pXLColumn[36];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[36]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[36]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[36]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[37]
                vXLIndexColumn = pXLColumn[37];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[37]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[37]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[37] = mSumValueColumn[37] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[37]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[37]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[38]
                vXLIndexColumn = pXLColumn[38];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[38]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[38]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[38] = mSumValueColumn[38] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[38]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[38]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[39]
                vXLIndexColumn = pXLColumn[39];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[39]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[39]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[39] = mSumValueColumn[39] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[39]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[39]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[40]
                vXLIndexColumn = pXLColumn[40];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[40]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[40]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[40] = mSumValueColumn[40] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[40]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[40]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[41]
                vXLIndexColumn = pXLColumn[41];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[41]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[41]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[41] = mSumValueColumn[41] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[41]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[41]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[42]
                vXLIndexColumn = pXLColumn[42];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[42]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[42] = mSumValueColumn[42] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[42]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[42]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[43]
                vXLIndexColumn = pXLColumn[43];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[43]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[43] = mSumValueColumn[43] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[43]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[43]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[44]
                vXLIndexColumn = pXLColumn[44];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[44]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[44] = mSumValueColumn[44] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[44]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[44]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[45]
                vXLIndexColumn = pXLColumn[45];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[45]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[45] = mSumValueColumn[45] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[45]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[45]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[46]
                vXLIndexColumn = pXLColumn[46];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[46]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[46] = mSumValueColumn[46] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[46]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[46]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[47]
                vXLIndexColumn = pXLColumn[47];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[47]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[47] = mSumValueColumn[47] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[47]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[47]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[48]
                vXLIndexColumn = pXLColumn[48];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[48]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[48] = mSumValueColumn[48] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[48]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[48]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[49]
                vXLIndexColumn = pXLColumn[49];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[49]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[49] = mSumValueColumn[49] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[49]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[49]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[50]
                vXLIndexColumn = pXLColumn[50];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[50]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[50] = mSumValueColumn[50] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[50]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[50]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[51]
                vXLIndexColumn = pXLColumn[51];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[51]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[51] = mSumValueColumn[51] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[51]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[51]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------

                //[52]
                vXLIndexColumn = pXLColumn[52];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[52]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[52]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[52]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[53]
                vXLIndexColumn = pXLColumn[53];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[53]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[53]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[53]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[54]
                vXLIndexColumn = pXLColumn[54];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[54]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[54]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[54] = mSumValueColumn[54] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[54]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[54]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[55]
                vXLIndexColumn = pXLColumn[55];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[55]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[55]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[55] = mSumValueColumn[55] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[55]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[55]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[56]
                vXLIndexColumn = pXLColumn[56];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[56]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[56]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[56] = mSumValueColumn[56] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[56]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[56]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[57]
                vXLIndexColumn = pXLColumn[57];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[57]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[57]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[57] = mSumValueColumn[57] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[57]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[57]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[58]
                vXLIndexColumn = pXLColumn[58];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[58]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertString(vGetValue, out vConvertString);
                    if (IsConvert == true)
                    {
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[58]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }

                    //vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    //IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    //if (IsConvert == true)
                    //{
                    //    mSumValueColumn[58] = mSumValueColumn[58] + vConvertDecimal;

                    //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                    //else
                    //{
                    //    if (vIsValueViewTemp == false)
                    //    {
                    //        vConvertString = string.Empty;
                    //    }
                    //    else
                    //    {
                    //        vConvertString = "[58]";
                    //    }
                    //    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    //}
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[58]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[59]
                vXLIndexColumn = pXLColumn[59];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[59]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[59] = mSumValueColumn[59] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[59]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[59]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[60]
                vXLIndexColumn = pXLColumn[60];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[60]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[60] = mSumValueColumn[60] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[60]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                    {
                        vConvertString = string.Empty;
                    }
                    else
                    {
                        vConvertString = "[60]";
                    }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[61]
                vXLIndexColumn = pXLColumn[61];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[61]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[61] = mSumValueColumn[61] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[61]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[61]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[62]
                vXLIndexColumn = pXLColumn[62];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[62]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[62] = mSumValueColumn[62] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[62]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[62]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[63]
                vXLIndexColumn = pXLColumn[63];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[63]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[63] = mSumValueColumn[63] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[63]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[63]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[64]
                vXLIndexColumn = pXLColumn[64];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[64]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[64] = mSumValueColumn[64] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[64]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[64]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[65]
                vXLIndexColumn = pXLColumn[65];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[65]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[65] = mSumValueColumn[65] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[65]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[65]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[66]
                vXLIndexColumn = pXLColumn[66];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[66]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[66] = mSumValueColumn[66] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[66]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[66]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[67]
                vXLIndexColumn = pXLColumn[67];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[67]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[67] = mSumValueColumn[67] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[67]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[67]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                //[68]
                vXLIndexColumn = pXLColumn[68];
                vGridIndexColumn = pGrid.GetColumnToIndex(pGridColumn[68]);
                //vGridIndexColumn = -1;
                if (vGridIndexColumn != -1)
                {
                    vGetValue = pGrid.GetCellValue(pRow, vGridIndexColumn);
                    IsConvert = IsConvertNumber(vGetValue, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        mSumValueColumn[68] = mSumValueColumn[68] + vConvertDecimal;

                        vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                    else
                    {
                        if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[68]";
                        }
                        mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                    }
                }
                else
                {
                    if (vIsValueViewTemp == false)
                        {
                            vConvertString = string.Empty;
                        }
                        else
                        {
                            vConvertString = "[68]";
                        }
                    mPrinting.XLSetCell(vXLine, vXLIndexColumn, vConvertString);
                }

                vXLine++;
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterfaceAdv.OnAppMessage(mMessageError);
            }


            pPrintingLine = vXLine;
            IsNewPage(pPrintingLine);
            if (mIsNewPage == true)
            {
                pPrintingLine = mPositionPrintLineSTART;
            }

            return pPrintingLine;
        }

        #endregion;

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPrintingLine)
        {
            if (mPrintingLineMAX < pPrintingLine)
            {
                mIsNewPage = true;
                mCopySumPrintingLine = CopyAndPaste(mPrinting, mCopySumPrintingLine);

                XlAllLineClear(mXLColumn);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int XLWirte(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pTerritory, string pUserName, string pCorporationName, string pYYYYMM, string pWageTypeName, string pDepartmentName)
        {
            string vMessage = string.Empty;
            mIsNewPage = false;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            mPringingDateTime = string.Format("{0} {1}", vPrintingDate, vPrintingTime);

            int vPrintingLine = mPositionPrintLineSTART;

            object vObject = null;
            int vGridIndexColumn = 0;
            string vDepartment = string.Empty;

            try
            {
                int vTotalRow = pGrid.RowCount;
                mPageTotalNumber = vTotalRow / 8;
                mPageTotalNumber = (vTotalRow % 8) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);

                int vCountRow = 0;
                if (vTotalRow > 0)
                {
                    SetArray(out mGridColumn, out mXLColumn);

                    mCorporationName = pCorporationName;   //��ü
                    mUserName = pUserName;                 //�����
                    mYYYYMM = pYYYYMM;                     //��³��
                    mWageTypeName = pWageTypeName;         //�޿�����
                    mDepartmentName = pDepartmentName;     //�μ�

                    vObject = pGrid.GetCellValue(mGridColumn[1]);
                    mDepartment = ConvertString(vObject);

                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vCountRow++;

                        vMessage = string.Format("Row : {0} / {1}", vRow, vTotalRow);
                        mAppInterfaceAdv.OnAppMessage(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        vGridIndexColumn = pGrid.GetColumnToIndex(mGridColumn[1]);
                        vObject = pGrid.GetCellValue(vRow, vGridIndexColumn);
                        vDepartment = ConvertString(vObject);

                        //////[�μ��հ�]
                        //////if (mDepartment != vDepartment)
                        //////{
                        //////    vPrintingLine = XLSUM(vPrintingLine, mXLColumn, mDepartment, mSumValueColumn);
                        //////    ClearValueSumValue();
                        //////    mDepartment = vDepartment;
                        //////}

                        vPrintingLine = XlLine(pGrid, vRow, vPrintingLine, mGridColumn, mXLColumn);

                        if (vTotalRow == vCountRow)
                        {
                            if (mPositionPrintLineSTART != vPrintingLine)
                            {
                                mCopySumPrintingLine = CopyAndPaste(mPrinting, mCopySumPrintingLine);
                            }

                            XlAllLineClear(mXLColumn);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mCountPage;
        }

        #endregion;

        #region ----- Excel Copy&Paste Methods ----

        //[Sheet2]������ [Sheet1]�� �ٿ��ֱ�
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine)
        {
            int vPrintHeaderColumnSTART = 1; //����Ǿ��� ��Ʈ�� ��, ���ۿ�
            int vPrintHeaderColumnEND = 63;  //����Ǿ��� ��Ʈ�� ��, ���῭

            mCountPage++;
            mPageString = string.Format("{0} / {1}", mCountPage, mPageTotalNumber);
            HeaderWrite(mUserName, mPringingDateTime, mYYYYMM, mWageTypeName, mDepartmentName, mPageString, mCorporationName);

            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;
            mPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = mPrinting.XLGetRange(vPrintHeaderColumnSTART, 1, mIncrementCopyMAX, vPrintHeaderColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            mPrinting.XLActiveSheet("Destination");
            object vRangeDestination = mPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, vPrintHeaderColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            mPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        #endregion;

        #region ----- Save Methods ----

        public void Save(string pSaveFileName)
        {
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D2}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(vSaveFileName);
        }

        #endregion;
    }
}