using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0603
{
    public class XLPrinting1
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        private int mPageNumber = 0;

        private string mXLOpenFileName = string.Empty;

        private int mPrintingLineSTART = 1;  //Line

        private int mCopyLineSUM = 1;        //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ, ���� �� ����
        private int mIncrementCopyMAX = 62; //����Ǿ��� ���� ����

        private int mCopyColumnSTART = 1;    //����Ǿ�  �� �� ���� ��
        private int mCopyColumnEND = 43;     //������ ���õ� ��Ʈ�� ����Ǿ��� �� �� ��ġ

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

        public XLPrinting1(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
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

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[82];
            pXLColumn = new int[82];

            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[0] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DEPT_NAME");                // �μ�                                                 
            pGDColumn[1] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PERSON_NUM");               // ���                                                 
            pGDColumn[2] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("NAME");                     // ����                                                 
            pGDColumn[3] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REPRE_NUM");                // �ֹι�ȣ                                             
            pGDColumn[4] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("JOIN_DATE");                // �Ի�����                                             
            pGDColumn[5] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("EXPIRE_DATE");              // �����߰�������                                       
            pGDColumn[6] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_DATE");              // ��������                                             
            pGDColumn[7] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("CONTINUE_DAY");             // �ٹ��ϼ�                                             
            pGDColumn[8] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ADDRESS");                  // �ּ�                                                 
            pGDColumn[9] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_DAY");                 // �����ϼ�                                             
            pGDColumn[10] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DED_DAY");                 // �����ϼ�                                             
            pGDColumn[11] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_NAME");             // ��������                                             
            pGDColumn[12] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ADJUSTMENT_TYPE");         // ���걸��(R : ������, N : �ߵ�����)                   
            pGDColumn[13] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE4");             // ���� ��¥1                                           
            pGDColumn[14] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE3");             // ���� ��¥2                                           
            pGDColumn[15] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE2");             // ���� ��¥3                                           
            pGDColumn[16] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE1");             // ���� ��¥4                                           
            pGDColumn[17] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE4");               // ������ ��¥1                                         
            pGDColumn[18] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE3");               // ������ ��¥2                                         
            pGDColumn[19] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE2");               // ������ ��¥3                                         
            pGDColumn[20] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE1");               // ������ ��¥4                                         
            pGDColumn[21] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_AMOUNT4");           // �޿�1                                                
            pGDColumn[22] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_AMOUNT3");           // �޿�2                                                
            pGDColumn[23] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_AMOUNT2");           // �޿�3                                                
            pGDColumn[24] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_AMOUNT1");           // �޿�4                                                
            pGDColumn[25] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_PAY_AMOUNT");        // �޿�(�հ�)                                           
            pGDColumn[26] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY4");                    // �ٹ��ϼ�1                                            
            pGDColumn[27] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY3");                    // �ٹ��ϼ�2                                            
            pGDColumn[28] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY2");                    // �ٹ��ϼ�3                                            
            pGDColumn[29] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY1");                    // �ٹ��ϼ�4                                            
            pGDColumn[30] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY_3RD_COUNT");           // �ٹ��ϼ�(�հ�)                                       
            pGDColumn[31] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_BOUNS");             // ���� �� ��                                         
            pGDColumn[32] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_BONUS_AMOUNT");      // ��(�հ�)                                           
            pGDColumn[33] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("YEAR_ALLOWANCE");          // ���� ������ ��                                       
            pGDColumn[34] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("YEAR_ALLOWANCE_AMOUNT");   // ������(�հ�)                                         
            pGDColumn[35] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_PAYMENT");           // �ӱ� �Ѿ�                                            
            pGDColumn[36] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY_AVR_PAY");             // �� ��� �ӱ�                                         
            pGDColumn[37] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_AMOUNT");           // �����޿�                                             
            pGDColumn[38] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_HONORARY_AMOUNT_1"); // ���������� ��                                      
            pGDColumn[39] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_INSUR_AMOUNT");     // ��������� ��                                        
            pGDColumn[40] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PRE_CORP_NAME");           // ����ó��-�����ٹ�������                              
            pGDColumn[41] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PRE_RETIRE_AMOUNT");       // �����޿���-�����ٹ�������                            
            pGDColumn[42] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PRE_INCOME_AMOUNT");       // �ҵ漼-�����ٹ�������                                
            pGDColumn[43] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PRE_LOCAL_AMOUNT");        // �ֹμ�-�����ٹ�������                                
            pGDColumn[44] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_PERIOD");             // �ټӱⰣ-�����ٹ�������                              
            pGDColumn[45] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_MONTH");              // �ټӿ���-�����ٹ�������                              
            pGDColumn[46] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DUPL_MONTH");              // �ߺ�����-�����ٹ�������                              
            pGDColumn[47] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_TOTAL_AMOUNT");     // �����޿���(���� ���� �޿�)                           
            pGDColumn[48] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_HONORARY_AMOUNT_2"); // �����޿���(�����̿� �����޿�)                        
            pGDColumn[49] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_DED_RATE1");        // �����޿�����(���������޿� �� �����̿� �����޿�)      
            pGDColumn[50] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_DED_AMOUNT");       // �����ҵ���� - �����޿����� - ���������޿�           
            pGDColumn[51] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_DED_RATE2");        // �����޿�����(���������޿� �� �����̿� �����޿�)      
            pGDColumn[52] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_INCOME_DED_AMOUNT");     // �����ҵ���� - �����޿����� - �����̿� �����޿�      
            pGDColumn[53] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_YEAR");               // �ټӿ�������(���������޿�)                           
            pGDColumn[54] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("LONG_DED_AMOUNT");         // �����ҵ���� - �ټӿ������� - ���������޿�           
            pGDColumn[55] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_LONG_YEAR");             // �ټӿ�������(�����̿� �����޿�)                      
            pGDColumn[56] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_LONG_DED_AMOUNT");       // �����ҵ���� - �ټӿ������� - �����̿� �����޿�      
            pGDColumn[57] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DED_SUM_AMOUNT");          // �����ҵ���� - �� - ���������޿�                     
            pGDColumn[58] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_DED_SUM_AMOUNT");        // �����ҵ���� - �� - �����̿� �����޿�                
            pGDColumn[59] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TAX_STD_AMOUNT");          // ���װ��ٰ� - �����ҵ����ǥ�� - ���������޿�       
            pGDColumn[60] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_TAX_STD_AMOUNT");        // ���װ��ٰ� - �����ҵ����ǥ�� - �����̿� �����޿�  
            pGDColumn[61] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("AVG_TAX_STD_AMOUNT");      // ���װ��ٰ� - ����հ���ǥ�� - ���������޿�         
            pGDColumn[62] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_AVG_TAX_STD_AMOUNT");    // ���װ��ٰ� - ����հ���ǥ�� - �����̿� �����޿�    
            pGDColumn[63] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TAX_RATE");                // �����ҵ漼��                                         
            pGDColumn[64] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("AVG_COMP_TAX_AMOUNT");     // ���װ��ٰ� - ����ջ��⼼�� - ���������޿�         
            pGDColumn[65] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_AVG_COMP_TAX_AMOUNT");   // ���װ��ٰ� - ����ջ��⼼�� - �����̿� �����޿�    
            pGDColumn[66] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("COMP_TAX_AMOUNT");         // ���װ��ٰ� - ���⼼�� - ���������޿�               
            pGDColumn[67] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_COMP_TAX_AMOUNT");       // ���װ��ٰ� - ���⼼�� - �����̿� �����޿�          
            pGDColumn[68] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TAX_DED_AMOUNT");          // ���װ��ٰ� - ���װ���(�ܱ�����) - ���������޿�     
            pGDColumn[69] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_TAX_DED_AMOUNT");        // ���װ��ٰ� - ���װ���(�ܱ�����) - �����̿� �����޿�
            pGDColumn[70] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_TAX_AMOUNT");       // �������� - �����ҵ漼 - ���������޿�                 
            pGDColumn[71] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("H_INCOME_TAX_AMOUNT");     // �������� - �����ҵ漼 - �����̿� �����޿�            
            pGDColumn[72] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RESIDENT_TAX_AMOUNT");     // �������� - �����ֹμ�                                
            pGDColumn[73] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REAL_TOTAL_AMOUNT");       // ���� ���޾�                                          
            pGDColumn[74] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ETC_SUPP_AMOUNT");         // ���� �� �����׸�(��Ÿ���� ����)                      
            pGDColumn[75] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ETC_DED_AMOUNT");          // ���� �� �����׸�(��Ÿ����)                           
            pGDColumn[76] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_TOTAL_AMOUNT1");    // �������޾�                                           
            pGDColumn[77] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("BANK_NAME");               // ����                                                 
            pGDColumn[78] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("BANK_ACCOUNTS");           // ���¹�ȣ                                             
            pGDColumn[79] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("NAME1");                   // ������
           
            //---------------------------------------------------------------------------------------------------------------------
            pXLColumn[0]  =  6; // �μ�                                                 
            pXLColumn[1]  = 16; // ���                                                 
            pXLColumn[2]  = 26; // ����                                                 
            pXLColumn[3]  = 36; // �ֹι�ȣ                                             
            pXLColumn[4]  =  6; // �Ի�����                                             
            pXLColumn[5]  = 16; // �����߰�������                                       
            pXLColumn[6]  = 26; // ��������                                             
            pXLColumn[7]  = 36; // �ٹ��ϼ�                                             
            pXLColumn[8]  =  6; // �ּ�                                                 
            pXLColumn[9]  = 26; // �����ϼ�                                             
            pXLColumn[10] = 36; // �����ϼ�                                             
            pXLColumn[11] = 26; // ��������                                             
            pXLColumn[12] = 38; // ���걸��(R : ������, N : �ߵ�����)                   
            pXLColumn[13] = 16; // ���� ��¥1                                           
            pXLColumn[14] = 21; // ���� ��¥2                                           
            pXLColumn[15] = 26; // ���� ��¥3                                           
            pXLColumn[16] = 31; // ���� ��¥4                                           
            pXLColumn[17] = 16; // ������ ��¥1                                         
            pXLColumn[18] = 21; // ������ ��¥2                                         
            pXLColumn[19] = 26; // ������ ��¥3                                         
            pXLColumn[20] = 31; // ������ ��¥4                                         
            pXLColumn[21] = 16; // �޿�1                                                
            pXLColumn[22] = 21; // �޿�2                                                
            pXLColumn[23] = 26; // �޿�3                                                
            pXLColumn[24] = 31; // �޿�4                                                
            pXLColumn[25] = 36; // �޿�(�հ�)                                           
            pXLColumn[26] = 16; // �ٹ��ϼ�1                                            
            pXLColumn[27] = 21; // �ٹ��ϼ�2                                            
            pXLColumn[28] = 26; // �ٹ��ϼ�3                                            
            pXLColumn[29] = 31; // �ٹ��ϼ�4                                            
            pXLColumn[30] = 36; // �ٹ��ϼ�(�հ�)                                       
            pXLColumn[31] = 16; // ���� �� ��                                         
            pXLColumn[32] = 36; // ��(�հ�)                                           
            pXLColumn[33] = 16; // ���� ������ ��                                       
            pXLColumn[34] = 36; // ������(�հ�)                                         
            pXLColumn[35] = 36; // �ӱ� �Ѿ�                                            
            pXLColumn[36] = 36; // �� ��� �ӱ�                                         
            pXLColumn[37] = 36; // �����޿�                                             
            pXLColumn[38] = 36; // ���������� ��                                      
            pXLColumn[39] = 36; // ��������� ��                                        
            pXLColumn[40] =  1; // ����ó��-�����ٹ�������                              
            pXLColumn[41] = 10; // �����޿���-�����ٹ�������                            
            pXLColumn[42] = 16; // �ҵ漼-�����ٹ�������                                
            pXLColumn[43] = 21; // �ֹμ�-�����ٹ�������                                
            pXLColumn[44] = 26; // �ټӱⰣ-�����ٹ�������                              
            pXLColumn[45] = 36; // �ټӿ���-�����ٹ�������                              
            pXLColumn[46] = 40; // �ߺ�����-�����ٹ�������                              
            pXLColumn[47] = 28; // �����޿���(���� ���� �޿�)                           
            pXLColumn[48] = 36; // �����޿���(�����̿� �����޿�)                        
            pXLColumn[49] = 28; // �����޿�����(���������޿� �� �����̿� �����޿�)      
            pXLColumn[50] = 31; // �����ҵ���� - �����޿����� - ���������޿�           
            pXLColumn[51] = 36; // �����޿�����(���������޿� �� �����̿� �����޿�)      
            pXLColumn[52] = 39; // �����ҵ���� - �����޿����� - �����̿� �����޿�      
            pXLColumn[53] = 28; // �ټӿ�������(���������޿�)                           
            pXLColumn[54] = 31; // �����ҵ���� - �ټӿ������� - ���������޿�           
            pXLColumn[55] = 36; // �ټӿ�������(�����̿� �����޿�)                      
            pXLColumn[56] = 39; // �����ҵ���� - �ټӿ������� - �����̿� �����޿�      
            pXLColumn[57] = 28; // �����ҵ���� - �� - ���������޿�                     
            pXLColumn[58] = 36; // �����ҵ���� - �� - �����̿� �����޿�                
            pXLColumn[59] = 28; // ���װ��ٰ� - �����ҵ����ǥ�� - ���������޿�       
            pXLColumn[60] = 36; // ���װ��ٰ� - �����ҵ����ǥ�� - �����̿� �����޿�  
            pXLColumn[61] = 28; // ���װ��ٰ� - ����հ���ǥ�� - ���������޿�         
            pXLColumn[62] = 36; // ���װ��ٰ� - ����հ���ǥ�� - �����̿� �����޿�    
            pXLColumn[63] = 28; // �����ҵ漼��                                         
            pXLColumn[64] = 28; // ���װ��ٰ� - ����ջ��⼼�� - ���������޿�         
            pXLColumn[65] = 36; // ���װ��ٰ� - ����ջ��⼼�� - �����̿� �����޿�    
            pXLColumn[66] = 28; // ���װ��ٰ� - ���⼼�� - ���������޿�               
            pXLColumn[67] = 36; // ���װ��ٰ� - ���⼼�� - �����̿� �����޿�          
            pXLColumn[68] = 28; // ���װ��ٰ� - ���װ���(�ܱ�����) - ���������޿�     
            pXLColumn[69] = 36; // ���װ��ٰ� - ���װ���(�ܱ�����) - �����̿� �����޿�
            pXLColumn[70] = 28; // �������� - �����ҵ漼 - ���������޿�                 
            pXLColumn[71] = 36; // �������� - �����ҵ漼 - �����̿� �����޿�            
            pXLColumn[72] = 28; // �������� - �����ֹμ�                                
            pXLColumn[73] = 28; // ���� ���޾�                                          
            pXLColumn[74] = 16; // ���� �� �����׸�(��Ÿ���� ����)                      
            pXLColumn[75] = 36; // ���� �� �����׸�(��Ÿ����)                           
            pXLColumn[76] = 36; // �������޾�                                           
            pXLColumn[77] =  6; // ����                                                 
            pXLColumn[78] = 21; // ���¹�ȣ                                             
            pXLColumn[79] = 33; // ������
        }

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_ALLOWANCE, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[90];
            pXLColumn = new int[90];

            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[0] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("PERSON_NUM");                  // ���                                              
            pGDColumn[1] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("OPERATING_UNIT_NAME");    //�����                 
            pGDColumn[2] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("JOIN_DATE");                      // �Ի���                            
            pGDColumn[3] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("CONTINUE_DAY");                    // �ټӱⰣ      
                                                   
            pGDColumn[4] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("NAME");                // ����                                             
            pGDColumn[5] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DEPT_NAME");              // �μ�
            pGDColumn[6] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_DATE_FR");              // ���������(�����)
            pGDColumn[7] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DED_DAY");          // �����Ⱓ

            pGDColumn[8] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REPRE_NUM");                  //�ֹι�ȣ
            pGDColumn[9] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ADDRESS");                 // �ּ�                                             
            pGDColumn[10] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_DATE");                 // �����                                             
            pGDColumn[11] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("BANK_ACCOUNTS");             // ���¹�ȣ

            pGDColumn[12] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("POST_NAME");         // ����                 
            pGDColumn[40] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("CLOSED_DATE");         // ������                 

            pGDColumn[13] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE4");             // ���� ��¥1                                           
            pGDColumn[14] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE3");             // ���� ��¥2                                           
            pGDColumn[15] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE2");             // ���� ��¥3                                           
            pGDColumn[16] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("START_DATE1");             // ���� ��¥4                                           
            pGDColumn[17] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE4");               // ������ ��¥1                                         
            pGDColumn[18] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE3");               // ������ ��¥2                                         
            pGDColumn[19] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE2");               // ������ ��¥3                                         
            pGDColumn[20] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("END_DATE1");               // ������ ��¥4                                             
            pGDColumn[26] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY4");                    // �ٹ��ϼ�1                               
            pGDColumn[27] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY3");                    // �ٹ��ϼ�2                               
            pGDColumn[28] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY2");                    // �ٹ��ϼ�3                               
            pGDColumn[29] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY1");                    // �ٹ��ϼ�4                               
            pGDColumn[30] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY_3RD_COUNT");           // �ٹ��ϼ�(�հ�)       
                                            
            pGDColumn[33] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_PAY_AMOUNT");        // �����޿� 
            pGDColumn[34] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("YEAR_BONUS_AMOUNT");     //������/����
            pGDColumn[35] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_PAYMENT");          //�հ�
            pGDColumn[36] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("DAY_AVR_PAY");                 // �� ��� �ӱ�                                         
            pGDColumn[37] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_AMOUNT");             // �����޿�                    
                                     
            pGDColumn[38] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_AMOUNT");            // ������                                   
            pGDColumn[39] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("GLORY_AMOUNT");            // ���α�                            
            pGDColumn[47] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_GROUP_AMOUNT");                  // ��ü���������  ?????
            pGDColumn[48] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("HONORARY_AMOUNT");       // �����������                     
            pGDColumn[49] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_PENSION_AMOUNT");                  //�������� ???
            pGDColumn[50] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_TAX_FREE_AMOUNT");                  //������ҵ�  ????        
            pGDColumn[51] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_AMOUNT");                  //�����Ѿ� ????

            pGDColumn[52] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("INCOME_TAX_AMOUNT_3");              //�����ҵ漼     
            pGDColumn[53] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RESIDENT_TAX_AMOUNT");             //��������ҵ漼
            pGDColumn[54] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("ETC_DED_AMOUNT");                     //��Ÿ���� 
            pGDColumn[55] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("TOTAL_DED_AMOUNT");           // �����Ѿ� 
            pGDColumn[56] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("RETIRE_TOTAL_AMOUNT1");            //�������޾� 
                                         
            pGDColumn[77] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("BANK_NAME");            // ����                                                   
            pGDColumn[79] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("NAME1");                   // ������

            ///////////////////////////LINE �κ�
            pGDColumn[80] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("RETIRE_SALARY_ITEM_NAME");  //�׸��                 
            pGDColumn[81] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT1");          //�������� -3     
            pGDColumn[82] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT2");          //�������� -2
            pGDColumn[83] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT3");          //�������� -1        
            pGDColumn[84] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT4");          //�������� 
            pGDColumn[85] = pGrid_PRINT_ALLOWANCE.GetColumnToIndex("AMOUNT");            //�ѱݾ�

            pGDColumn[86] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("SUB_INCOME_TAX_AMOUNT");              //�����ҵ漼     
            pGDColumn[87] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("SUB_RESIDENT_TAX_AMOUNT");             //��������ҵ漼
            pGDColumn[88] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REAL_TOTAL_DED_AMOUNT");             //��������ҵ漼
            pGDColumn[89] = pGrid_RETIRE_ADJUSTMENT.GetColumnToIndex("REAL_RETIRE_TOTAL_AMOUNT");            //�������޾� 


            //---------------------------------------------------------------------------------------------------------------------
            /////������ ��������
            pXLColumn[0] = 6; // ���                                               
            pXLColumn[1] = 6; // �����                                        
            pXLColumn[2] = 6; // �Ի���                                          
            pXLColumn[3] = 6; // �ټӱⰣ 
                                                 
            pXLColumn[4] = 18; // ����                                           
            pXLColumn[5] = 18; //�μ�
            pXLColumn[6] = 18; // �����
            pXLColumn[7] = 18; //�����Ⱓ 

            pXLColumn[8] = 28; // �ֹι�ȣ
            pXLColumn[9] = 28; // �ּ�                                          
            pXLColumn[10] = 28; // �����                                        
            pXLColumn[11] = 33; // ���¹�ȣ                         
                         
            pXLColumn[12] = 38; // ����    
          //  pXLColumn[12] = 38; // ������ 

            ///// ������ �������� ���
            pXLColumn[16] = 8; // ���� ��¥1
            pXLColumn[15] = 15; // ���� ��¥2                                           
            pXLColumn[14] = 22; // ���� ��¥3                                           
            pXLColumn[13] = 29; // ���� ��¥4              
                                         
            pXLColumn[20] = 8; // ������ ��¥1                                         
            pXLColumn[19] = 15;// ������ ��¥2                                         
            pXLColumn[18] = 22; // ������ ��¥3                                         
            pXLColumn[17] = 29; // ������ ��¥4           

            pXLColumn[29] = 8; // �ٹ��ϼ�1                                            
            pXLColumn[28] = 15; // �ٹ��ϼ�2                                            
            pXLColumn[27] = 22; // �ٹ��ϼ�3                                            
            pXLColumn[26] = 29; // �ٹ��ϼ�4                                            
            pXLColumn[30] = 36; // �ٹ��ϼ�(�հ�) 
                                                                        
            pXLColumn[33] = 1; // �����޿�                               
            pXLColumn[34] = 10; // ������+����
            pXLColumn[35] = 18; // �հ�                                     
            pXLColumn[36] = 28; // �� ��� �ӱ�                                         
            pXLColumn[37] = 36; // �����޿�           
                                              
            pXLColumn[38] = 13;  //���������޳��� �ݾ�    
            pXLColumn[39] = 35;  //�����ݰ������� �ݾ�

            pXLColumn[77] = 28;  //����� 

            //////LINE�κ�
            pXLColumn[80] = 1; // �޿��׸�
            pXLColumn[81] = 8; // �������� -3                                   
            pXLColumn[82] = 15; // �������� -2                                
            pXLColumn[83] = 22; // �������� -1                                        
            pXLColumn[84] = 29; // ��������                                         
            pXLColumn[85] = 36; // �հ� 
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

        #region ----- Excel Wirte [Header] Methods ----
        private int HeaderWrite(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, int pGridRow, int[] pGDColumn, int[] pXLColumn)
        {
            int vXLine = 8; // ������ ������ ǥ�õǴ� �� ��ȣ

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;


            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("SourceTab1");

                
                // ��� 
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����� 
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }

                // �Ի���
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, DBNull.Value);
                }

                // �ټӱⰣ 
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine+6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+6, vXLColumnIndex, vConvertString);
                }
                

                // ����
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    mPrinting.XLSetCell(80, 28, vConvertString);
                    mPrinting.XLSetCell(78, 26, DateTime.Today );
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �μ�
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }

                // �����
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, DBNull.Value);
                }

                // �����Ⱓ 
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                
                // �ֹι�ȣ
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ּ�
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }

                // �����
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, DBNull.Value);
                }

                // ����
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", "("+vConvertString+")");
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                // ���¹�ȣ
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine + 6, vXLColumnIndex, vConvertString);
                }

                // ����
                vGDColumnIndex = pGDColumn[12];  
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {                 
                    vConvertString = string.Format("{0}", vConvertString);
                    mPrinting.XLSetCell(vXLine , vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine , vXLColumnIndex, vConvertString);
                }

                // ������
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine + 4, vXLColumnIndex, DBNull.Value);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 13;
                //-------------------------------------------------------------------

                // �����޿� 
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // ������/����
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // �հ� 
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // �����  
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // �����޿� 
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ���� ��¥1 
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ���� ��¥2 
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ���� ��¥3 
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ���� ��¥4
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // ������ ��¥1
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ������ ��¥2
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ������ ��¥3
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ������ ��¥4
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }
                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // �ٹ��ϼ�1 
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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
                // �ٹ��ϼ�2
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ٹ��ϼ�3
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ٹ��ϼ�4
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ٹ��ϼ��հ�
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 31;
                //-------------------------------------------------------------------

                // ������  
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // �����ҵ漼  
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ���α� 
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // ��������ҵ� 
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ��ü���������  
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // ��Ÿ����
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // ���������� 
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+2, vXLColumnIndex, vConvertString);
                }
                // �������� 
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine+4, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+4, vXLColumnIndex, vConvertString);
                }
                // ������ҵ� 
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine+6, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine+6, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 8;
                //-------------------------------------------------------------------
                // �����Ѿ� 
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                // �����Ѿ�  
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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
                // �������޾�  
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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
            

            return vXLine;
        }

        #endregion

        #region ----- Line Write Method 1 -----
        private int XLLine1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_ETC_ALLOWANCE, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pCourse)
        {
            int vXLine = pXLine; // ������ ������ ǥ�õǴ� �� ��ȣ

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            int vGDColumnIndex1 = 0;
            int vXLColumnIndex1 = 0;
            int vGDColumnIndex2 = 0;
            int vXLColumnIndex2 = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                int vCountRow = pGrid_ETC_ALLOWANCE.RowCount; //pGrid_ETC_ALLOWANCE �׸����� �� ���

                mPrinting.XLActiveSheet("Destination");

                //-------------------------------------------------------------------
                vXLine = vXLine + 5;
                //-------------------------------------------------------------------

                // �μ�
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ֹε�Ϲ�ȣ
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �Ի�����
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // �����߰�������
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ��������
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // �ٹ��ϼ�
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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
                // �ּ�
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����ϼ�
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ϼ�
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
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

                // ��������
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���걸��
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                    if (vConvertString == "R")
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, "��");
                    }
                    else
                    {
                        mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 4), "��");
                    }
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    mPrinting.XLSetCell(vXLine, (vXLColumnIndex + 4), vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // ���� ��¥1 
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ���� ��¥2 
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ���� ��¥3 
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ���� ��¥4
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // ������ ��¥1
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ������ ��¥2
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ������ ��¥3
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                // ������ ��¥4
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertDate(vObject, out vConvertDateTime);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDateTime.ToShortDateString());
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, DBNull.Value);
                }

                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------

                // �޿�1
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �޿�2
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �޿�3
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �޿�4
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �޿�(�հ�) 
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �ٹ��ϼ�1
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ٹ��ϼ�2
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ٹ��ϼ�3
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ٹ��ϼ�4
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ٹ��ϼ�(�հ�)
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,##0}", vConvertDecimal);
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

                // ���� �� ��
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��(�հ�)
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ���� ������ ��
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������(�հ�) 
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �ӱ� �Ѿ�
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �� ��� �ӱ�  
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �����޿�
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ���������� ��
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ��������� �� 
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ����ó��-�����ٹ�������
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����޿���-�����ٹ�������
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ҵ漼-�����ٹ�������
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ֹμ�-�����ٹ�������
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӱⰣ-�����ٹ������� 
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ���-�����ٹ�������
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ߺ�����-�����ٹ�������  
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����޿���(���� ���� �޿�)    
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿���(�����̿� �����޿�) 
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �����޿�����(���������޿� �� �����̿� �����޿�) 
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����ҵ���� - �����޿����� - ���������޿� 
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿�����(���������޿� �� �����̿� �����޿�)
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����ҵ���� - �����޿����� - �����̿� �����޿�
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �ټӿ�������(���������޿�)
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ҵ���� - �ټӿ������� - ���������޿�
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ټӿ�������(�����̿� �����޿�)
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ҵ���� - �ټӿ������� - �����̿� �����޿�
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �����ҵ���� - �� - ���������޿� 
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ҵ���� - �� - �����̿� �����޿�
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ���װ��ٰ� - �����ҵ����ǥ�� - ���������޿�
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���װ��ٰ� - �����ҵ����ǥ�� - �����̿� �����޿�
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ���װ��ٰ� - ����հ���ǥ�� - ���������޿�
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���װ��ٰ� - ����հ���ǥ�� - �����̿� �����޿�
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �����ҵ漼��
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���װ��ٰ� - ����ջ��⼼�� - ���������޿�
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���װ��ٰ� - ����ջ��⼼�� - �����̿� �����޿�
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ���װ��ٰ� - ���⼼�� - ���������޿�    
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���װ��ٰ� - ���⼼�� - �����̿� �����޿�
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                //  ���װ��ٰ� - ���װ���(�ܱ�����) - ���������޿�     
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���װ��ٰ� - ���װ���(�ܱ�����) - �����̿� �����޿�
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �������� - �����ҵ漼 - ���������޿�   
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� - �����ҵ漼 - �����̿� �����޿�   
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // �������� - �����ֹμ�
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

                // ���� ���޾�
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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



                if (vCountRow > 0)
                {
                    for (int vRow = 0; vRow < vCountRow; vRow++)
                    {
                        vGDColumnIndex1 = pGrid_ETC_ALLOWANCE.GetColumnToIndex("ALLOWANCE_DESC");
                        vXLColumnIndex1 = 16;

                        vGDColumnIndex2 = pGrid_ETC_ALLOWANCE.GetColumnToIndex("ALLOWANCE_AMOUNT");
                        vXLColumnIndex2 = 36;

                        vObject = pGrid_ETC_ALLOWANCE.GetCellValue(vRow, vGDColumnIndex1);
                        IsConvert = IsConvertString(vObject, out vConvertString);
                        if (IsConvert == true)
                        {
                            vConvertString = string.Format("{0}", vConvertString);
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex1, vConvertString);
                        }
                        else
                        {
                            vConvertString = string.Empty;
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex1, vConvertString);
                        }

                        vObject = pGrid_ETC_ALLOWANCE.GetCellValue(vRow, vGDColumnIndex2);
                        IsConvert = IsConvertString(vObject, out vConvertString);
                        IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                        if (IsConvert == true)
                        {
                            vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex2, vConvertString);
                        }
                        else
                        {
                            vConvertString = string.Empty;
                            mPrinting.XLSetCell(vXLine, vXLColumnIndex2, vConvertString);
                        }
                        vXLine = vXLine + 1;
                    }
                }

                // ���� �� �����׸�(��Ÿ���� ����)    
                //vGDColumnIndex = pGDColumn[74];
                //vXLColumnIndex = pXLColumn[74];
                //vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertString(vObject, out vConvertString);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0}", vConvertString);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                // ���� �� �����׸�(��Ÿ����)  
                //vGDColumnIndex = pGDColumn[75];
                //vXLColumnIndex = pXLColumn[75];
                //vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                //-------------------------------------------------------------------
                vXLine = 50;
                //-------------------------------------------------------------------

                // �������޾�
                vGDColumnIndex = pGDColumn[76];
                vXLColumnIndex = pXLColumn[76];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //-------------------------------------------------------------------
                vXLine = 53;
                //-------------------------------------------------------------------

                // ����
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���¹�ȣ
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��¥
                vConvertString = string.Format("{0}", iDate.ISGetDate().ToShortDateString());
                mPrinting.XLSetCell(vXLine, 34, vConvertString);

                //-------------------------------------------------------------------
                vXLine = vXLine + 4;
                //-------------------------------------------------------------------

                // ������
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_RETIRE_ADJUSTMENT.GetCellValue(pGridRow, vGDColumnIndex);

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

        private int XLLine2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_ETC_ALLOWANCE, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pCourse)
        {
            int vXLine = pXLine; // ������ ������ ǥ�õǴ� �� ��ȣ

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;
            

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                int vCountRow = pGrid_ETC_ALLOWANCE.RowCount; //pGrid_ETC_ALLOWANCE �׸����� �� ���

                mPrinting.XLActiveSheet("Destination");

               
                // �׸� 
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

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



                // �������� -3 
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� -2
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� -1
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ��������
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //(�հ�) 
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_ETC_ALLOWANCE.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###,###,###,###,###}", vConvertDecimal);
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

        #region ----- Excel Write RetireAdjustment  Method ----

        public int WriteRetireAdjustment(string pPrint_Type, string pSaveFileName, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_ETC_ALLOWANCE)
        {
            
            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn;
            int[] vXLColumn;

            int vTotalRow = pGrid_RETIRE_ADJUSTMENT.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 0;

            int vSecondPrinting = 9; //1�δ� 3�������̹Ƿ�, 3*10=30��°�� �μ�
            int vCountPrinting = 0;

            SetArray1(pGrid_RETIRE_ADJUSTMENT, out vGDColumn, out vXLColumn);

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;
                pGrid_RETIRE_ADJUSTMENT.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                if (isOpen == true)
                {
                    vCountPrinting++;

                    mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                    vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                    pGrid_RETIRE_ADJUSTMENT.CurrentCellMoveTo(vRow, 0);
                    pGrid_RETIRE_ADJUSTMENT.Focus();
                    pGrid_RETIRE_ADJUSTMENT.CurrentCellActivate(vRow, 0);

                    // ������ ���� ����
                    vPrintingLine = XLLine1(pGrid_RETIRE_ADJUSTMENT, pGrid_ETC_ALLOWANCE, vRow, vPrintingLine, vGDColumn, vXLColumn, "SRC_TAB1");

                    if (vSecondPrinting < vCountPrinting)
                    {
                        if (pPrint_Type == "FILE")
                        {
                            ////���� ����
                            SAVE(pSaveFileName);
                        }
                        else
                        {
                            Printing(1, vSecondPrinting);
                        }

                        mPrinting.XLOpenFileClose();
                        isOpen = XLFileOpen();

                        vCountPrinting = 0;
                        vPrintingLine = 1;
                        mCopyLineSUM = 1;
                    }
                    else if (vTotalRow == vRowCount)
                    {
                        if (pPrint_Type == "FILE")
                        {
                            ////���� ����
                            SAVE(pSaveFileName);
                        }
                        else
                        {
                            Printing(1, vCountPrinting);
                            //PreView(1, vCountPrinting);       
                        }
                    }
                }
            }
            
            mPrinting.XLOpenFileClose();

            return mPageNumber;
        }

        public int WriteRetireAdjustment_NFK(string pPrint_Type, string pSaveFileName, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_RETIRE_ADJUSTMENT, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_ETC_ALLOWANCE)
        {
            mCopyLineSUM = 1;        //������ ���õ� ��Ʈ�� ����Ǿ��� ���� �� ��ġ, ���� �� ����
            mIncrementCopyMAX = 85; //����Ǿ��� ���� ����
           
            mCopyColumnSTART = 1;    //����Ǿ�  �� �� ���� ��
            mCopyColumnEND = 43;     //������ ���õ� ��Ʈ�� ����Ǿ��� �� �� ��ġ

            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn;
            int[] vXLColumn;

            int vTotalRow = pGrid_ETC_ALLOWANCE.RowCount;
            int vH_Row = 0;//  pGrid_RETIRE_ADJUSTMENT.RowCount;
            int vRowCount = 0;

            int vPrintingLine = 30;

            int vSecondPrinting = 11; //1�δ� 3�������̹Ƿ�, 3*10=30��°�� �μ�
            int vCountPrinting = 0;

            SetArray2(pGrid_RETIRE_ADJUSTMENT, pGrid_ETC_ALLOWANCE, out vGDColumn, out vXLColumn);

            HeaderWrite(pGrid_RETIRE_ADJUSTMENT, vH_Row, vGDColumn, vXLColumn);

            mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");

            for (int vRow = 0; vRow < vTotalRow; vRow++)
            {
                vRowCount++;
                pGrid_ETC_ALLOWANCE.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                mAppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                if (isOpen == true)
                {
                    vCountPrinting++;

                    pGrid_ETC_ALLOWANCE.CurrentCellMoveTo(vRow, 0);
                    pGrid_ETC_ALLOWANCE.Focus();
                    pGrid_ETC_ALLOWANCE.CurrentCellActivate(vRow, 0);

                    // ������ ���� ����
                    vPrintingLine = XLLine2(pGrid_ETC_ALLOWANCE, vRow, vPrintingLine, vGDColumn, vXLColumn, "SRC_TAB1");

                    if (vTotalRow == vRowCount)
                    {
                        if (pPrint_Type == "FILE")
                        {
                            DeleteSheet();
                            ////���� ����
                            SAVE(pSaveFileName);
                        }
                        else
                        {
                            DeleteSheet();
                            Printing(1, vCountPrinting);
                            //PreView(1, vCountPrinting);       
                        }
                    }
                }
            }

            mPrinting.XLOpenFileClose();

            return mPageNumber;
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //ù��° ������ ����
        private int CopyAndPaste(XL.XLPrint pPrinting, int pCopySumPrintingLine, string pCourse)
        {
            int vCopySumPrintingLine = pCopySumPrintingLine;

            int vCopyPrintingRowSTART = vCopySumPrintingLine;
            vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            int vCopyPrintingRowEnd = vCopySumPrintingLine;

 
            pPrinting.XLActiveSheet("SourceTab1");
            object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet("Destination");
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            mPageNumber++; //������ ��ȣ

            return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPrinting(pPageSTART, pPageEND);
        }

        public void DeleteSheet()
        {
            mPrinting.XLDeleteSheet("SourceTab1");
        }

        public void PreView(int pPageSTART, int pPageEND)
        {
            try
            {
                mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }
        }

        #endregion;


        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder.ToString(), vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
        
        #endregion;
    }
}