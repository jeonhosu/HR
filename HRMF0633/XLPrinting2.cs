using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0633
{
    public class XLPrinting2
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
        private int mIncrementCopyMAX = 71; //����Ǿ��� ���� ����

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

        public XLPrinting2(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
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

        #region ----- SetArray1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[162];
            pXLColumn = new int[162];

            //--------------------------------------------------------------------------------------------------------------------
            pGDColumn[0]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TYPE_1");           // ���� ����(������1)       
            pGDColumn[1]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TYPE_2");           // ���� ����(������2)       
            pGDColumn[2]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONALITY_TYPE_1");        // ���ܱ��� ����(������1)   
            pGDColumn[3]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATIONALITY_TYPE_9");        // ���ܱ��� ����(�ܱ���9)      
            pGDColumn[4]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATION_DESC");               // ��������
            pGDColumn[5]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NATION_ISO_CODE");           // ���������ڵ�
            pGDColumn[6]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OWNER_TYPE");                // ¡���ǹ��ڱ���
            pGDColumn[7]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("VAT_NUMBER");                // ����ڵ�Ϲ�ȣ                                       
            pGDColumn[8]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("CORP_NAME");                 // ���θ�(��ȣ)                                         
            pGDColumn[9]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRESIDENT_NAME");            // ��ǥ��(����)                                         
            pGDColumn[10]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ORG_ADDRESS");               // ������(�ּ�)   
                                      
            pGDColumn[11]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NAME");                      // ����                                                 
            pGDColumn[12]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPRE_NUM");                 // �ֹι�ȣ                                             
            pGDColumn[13]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADDRESS");                   // �ּ�                                                 
            pGDColumn[14]   = pGrid_WITHHOLDING_TAX.GetColumnToIndex("START_RETIRE_DATE");         // �ͼӿ��� ���� ����                                   
            pGDColumn[15]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_NAME");               // ��������                                             
            pGDColumn[16]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LAST_RETIRE_DATE");          // �ͼӿ��� ������ ����(��������)                       
            pGDColumn[17]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TAX1");               // �������װ�������1                                    
            pGDColumn[18]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TAX2");               // �������װ�������2                                    
            pGDColumn[19]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TAX3");               // �������װ�������3                                    
            pGDColumn[20]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TAX4");               // �������װ������� �հ�   
                             
            pGDColumn[21]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME1");           // �ٹ�ó��1                                            
            pGDColumn[22]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME2");           // �ٹ�ó��2                                            
            pGDColumn[23]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME3");           // �ٹ�ó��3                                            
            pGDColumn[24]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_CORP_NAME4");           // �ٹ�ó�� �հ�                                        
            pGDColumn[25]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER1");          // ����ڵ�Ϲ�ȣ1                                      
            pGDColumn[26]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER2");          // ����ڵ�Ϲ�ȣ2                                      
            pGDColumn[27]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER3");          // ����ڵ�Ϲ�ȣ3                                      
            pGDColumn[28]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("WORK_VAT_NUMBER4");          // ����ڵ�Ϲ�ȣ �հ�                                  
            pGDColumn[29]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TOTAL_AMOUNT1");      // �����޿�1                                            
            pGDColumn[30]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TOTAL_AMOUNT2");      // �����޿�2          
                                  
            pGDColumn[31]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TOTAL_AMOUNT3");      // �����޿�3                                            
            pGDColumn[32]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_TOTAL_AMOUNT4");      // �����޿� �հ�               
            pGDColumn[33]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT1");          // ����������(�߰�������)1                            
            pGDColumn[34]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT2");          // ����������(�߰�������)2                            
            pGDColumn[35]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT3");          // ����������(�߰�������)3                            
            pGDColumn[36]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT4");          // ����������(�߰�������) �հ�                        
            pGDColumn[37]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT1");            // ���������Ͻñ�1                                      
            pGDColumn[38]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT2");            // ���������Ͻñ�2                                      
            pGDColumn[39]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT3");            // ���������Ͻñ�3                                      
            pGDColumn[40]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT4");            // ���������Ͻñ� �հ�   

            pGDColumn[41]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_HONORARY_AMOUNT1");   // ���������� ���������Ͻñ�  
            pGDColumn[42]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_HONORARY_AMOUNT2");   // ���������� ���������Ͻñ�  
            pGDColumn[43]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_HONORARY_AMOUNT3");   // ���������� ���������Ͻñ�  
            pGDColumn[44]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_HONORARY_AMOUNT4");   // ���������� ���������Ͻñ� �հ�  
            pGDColumn[45]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL1");                    // ������1                                                  
            pGDColumn[46]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL2");                    // ������2                                                  
            pGDColumn[47]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL3");                    // ������3                                                  
            pGDColumn[48]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL4");                    // ������ �հ� 
            pGDColumn[49] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_HONORARY_AMOUNT1");     // �����ܰ�1                                                  
            pGDColumn[50] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_HONORARY_AMOUNT2");     // �����ܰ�2  
                                                
            pGDColumn[51] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_HONORARY_AMOUNT3");     // �����ܰ�3                                                  
            pGDColumn[52] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_HONORARY_AMOUNT4");     // �����ܰ� �հ� 
            pGDColumn[53]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NON_TAX1");                  // ������ҵ�1                                          
            pGDColumn[54]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NON_TAX2");                  // ������ҵ�2                                          
            pGDColumn[55]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NON_TAX3");                  // ������ҵ�3                                          
            pGDColumn[56]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("NON_TAX4");                  // ������ҵ� �հ�                                      
            pGDColumn[57]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_RECEIPTS1");           // �Ѽ��ɾ�1                                            
            pGDColumn[58]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPAY_TOTAL_AMOUNT1");       // ������ �հ��1               
            pGDColumn[59]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_MONEY_DUE1");         // �ҵ��� ���Ծ�1                                       
            pGDColumn[60]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANNUITY_DED1");       // �������� �ҵ������1                                 

            pGDColumn[61]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANNUITY_LUMP_SUM1");  // ���������Ͻñ�1                                      
            pGDColumn[62]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_RECEIPTS2");           // �Ѽ��ɾ�2                                            
            pGDColumn[63]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("REPAY_TOTAL_AMOUNT2");       // ������ �հ��2                                       
            pGDColumn[64]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_MONEY_DUE2");         // �ҵ��� ���Ծ�2                                       
            pGDColumn[65]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANNUITY_DED2");       // �������� �ҵ������2                                 
            pGDColumn[66]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANNUITY_LUMP_SUM2");  // ���������Ͻñ�2                                      
            pGDColumn[67]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANN_LUMP_SUM_E1");    // ���������Ͻñ� ���޿����, �̿��ݾ�                  
            pGDColumn[68]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_LUMP_SUM1");           // ���Ͻñ�                                  
            pGDColumn[69]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECEIVE_RETIRE_PAY1");       // ���ɰ��������޿���                                   
            pGDColumn[70]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_RETIRE_ANN_DED1");        // ȯ�������ҵ����
                                     
            pGDColumn[71]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_RETIRE_ANN_STANDARD1");   // ȯ�������ҵ����ǥ��                                 
            pGDColumn[72]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_YEARLY_ANN_STANDARD1");   // ȯ�꿬��� ����ǥ��                                  
            pGDColumn[73]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_YEARLY_TAX_AMOUNT1");     // ȯ�� ����� ���⼼��                                 
            pGDColumn[74]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANN_LUMP_SUM_E2");    // ���������Ͻñ� ���޿����, �̿��ݾ�                  
            pGDColumn[75]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANN_LUMP_SUM_E3");    // ���������Ͻñ� ���޿����, �̿��ݾ�                  
            pGDColumn[76]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_LUMP_SUM2");           // ���Ͻñ�                                             
            pGDColumn[77]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RECEIVE_RETIRE_PAY2");       // ���ɰ��������޿���               
            pGDColumn[78]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_RETIRE_ANN_DED2");        // ȯ�������ҵ����          
            pGDColumn[79]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_RETIRE_ANN_STANDARD2");   // ȯ�������ҵ����ǥ��                                 
            pGDColumn[80]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_YEARLY_ANN_STANDARD2");   // ȯ�꿬��� ����ǥ��                                  

            pGDColumn[81]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EX_YEARLY_TAX_AMOUNT2");     // ȯ�� ����� ���⼼��                                 
            pGDColumn[82]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_ANN_LUMP_SUM_E4");    // ���������Ͻñ� ���޿����, �̿��ݾ�                  
            pGDColumn[83] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_DATE_FR1");           // �Ի���(���������)1     [ ���� �����޿�(��(��)�ٹ���) ]                             
            pGDColumn[84]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_DATE1");              // �����1                                              
            pGDColumn[85]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_MONTH1");               // �ټӿ���1                                            
            pGDColumn[86]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EXCEPT_MONTH1");             // ���ܿ���1                                            
            pGDColumn[87]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_YEAR1");                // �ټӿ���1                                            
            pGDColumn[88] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ORI_JOIN_DATE2");            // �Ի���2     [ ���� �� �����޿�(��(��)�ٹ���) ]                                        
            pGDColumn[89]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_DATE2");              // �����2                                              
            pGDColumn[90]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_MONTH2");               // �ټӿ���2                                            

            pGDColumn[91]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EXCEPT_MONTH2");             // ���ܿ���2                                            
            pGDColumn[92]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("LONG_YEAR2");                // �ټӿ���2                                            
            pGDColumn[93] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_RETIRE_DATE_FR1");       // �Ի���(���������)1    [ ���� �����޿�(��(��)�ٹ���) ]                              
            pGDColumn[94]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_RETIRE_DATE1");          // �����1                                              
            pGDColumn[95]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LONG_MONTH1");           // �ټӿ���1                                            
            pGDColumn[96]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_EXCEPT_MONTH1");         // ���ܿ���1                                            
            pGDColumn[97] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_ORI_JOIN_DATE2");        // �Ի���2                 [ ���� �� �����޿�(��(��)�ٹ���) ]                                          
            pGDColumn[98]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_RETIRE_DATE2");          // �����2                                              
            pGDColumn[99]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LONG_MONTH2");           // �ټӿ���2                                            
            pGDColumn[100]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_EXCEPT_MONTH2");         // ���ܿ���2  

            pGDColumn[101]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LONG_YEAR1");            // �ߺ�����1             [ ���� �����޿�(��(��)�ٹ���) ]                    
            pGDColumn[102]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LONG_YEAR2");            // �ߺ�����2             [ ���� �� �����޿�(��(��)�ٹ���) ]     
            pGDColumn[103]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_AMOUNT");             // �����޿���                                           
            pGDColumn[104]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("HONORARY_AMOUNT");           // ��������                                           
            pGDColumn[105]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_RETIRE_AMOUNT");       // �����޿��� �հ�                                      
            pGDColumn[106]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DED_SUM_AMOUNT");            // �����ҵ���� - ��                                    
            pGDColumn[107]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_DED_SUM_AMOUNT");          // �����ҵ���� - ��                                    
            pGDColumn[108]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_INCOME_DED_AMOUNT");   // �����ҵ���� �հ�                                    /
            pGDColumn[109]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_STD_AMOUNT");            // �����ҵ����ǥ�� - ���������޿�                      
            pGDColumn[110]  = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_TAX_STD_AMOUNT");          // �����ҵ����ǥ�� - �����̿� �����޿�                 

            pGDColumn[111] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_TAX_STD_AMOUNT");      // �����ҵ����ǥ�� �հ�.                              
            pGDColumn[112] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("AVG_TAX_STD_AMOUNT");        // ���װ��ٰ� - ����հ���ǥ�� - ���������޿�         
            pGDColumn[113] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_AVG_TAX_STD_AMOUNT");      // ���װ��ٰ� - ����հ���ǥ�� - �����̿� �����޿�    
            pGDColumn[114] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AVG_TAX_STD_AMOUNT");   // ����հ���ǥ�� �հ�.                   
            pGDColumn[115] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("AVG_COMP_TAX_AMOUNT");       // ���װ��ٰ� - ����ջ��⼼�� - ���������޿�         
            pGDColumn[116] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_AVG_COMP_TAX_AMOUNT");     // ���װ��ٰ� - ����ջ��⼼�� - �����̿� �����޿�    
            pGDColumn[117] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_AVG_COMP_TAX_AMOUNT"); // ����ջ��⼼�� �հ�                                  
            pGDColumn[118] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("COMP_TAX_AMOUNT");           // ���װ��ٰ� - ���⼼�� - ���������޿�               
            pGDColumn[119] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_COMP_TAX_AMOUNT");         // ���װ��ٰ� - ���⼼�� - �����̿� �����޿�          
            pGDColumn[120] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_COMP_TAX_AMOUNT");     // ���⼼�� �հ�                                        

            pGDColumn[121] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TAX_DED_AMOUNT");            // ���װ��ٰ� - ���װ���(�ܱ�����) - ���������޿�     
            pGDColumn[122] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_TAX_DED_AMOUNT");          // ���װ��ٰ� - ���װ���(�ܱ�����) - �����̿� �����޿�
            pGDColumn[123] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_TAX_DED_AMOUNT");      // ���װ��� �հ�                                        
            pGDColumn[124] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("INCOME_TAX_AMOUNT");         // �������� - �����ҵ漼 - ���������޿�                 
            pGDColumn[125] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("H_INCOME_TAX_AMOUNT");       // �������� - �����ҵ漼 - �����̿� �����޿�            
            pGDColumn[126] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_INCOME_TAX_AMOUNT");   // �������� �հ�                                        
            pGDColumn[127] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_INCOME_TAX_AMOUNT1");  // �������� - �ҵ漼 �հ�                               
            pGDColumn[128] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RESIDENT_TAX_AMOUNT");       // �������� - �ֹμ� �հ�                               
            pGDColumn[129] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_SP_TAX_AMOUNT");       // ����� Ư���� �հ�                                   
            pGDColumn[130] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_INCOME_TAX_AMOUNT2");  // �������� �հ�                                        

            pGDColumn[131] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_INCOME_AMOUNT");         // ��(��)�ٹ��� �ⳳ�μ��� | �ҵ漼                     
            pGDColumn[132] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_LOCAL_AMOUNT");          // ��(��)�ٹ��� �ⳳ�μ��� | �ֹμ�                     
            pGDColumn[133] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_SP_TAX_AMOUNT");         // ��(��)�ٹ��� �ⳳ�μ��� | ��Ư��                     
            pGDColumn[134] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_TOTAL");                 // ��(��)�ٹ��� �ⳳ�μ��� | �հ�                       
            pGDColumn[135] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MINUS_INCOME_AMOUNT");       // ������õ¡������ | �ҵ漼                            
            pGDColumn[136] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MINUS_LOCAL_AMOUNT");        // ������õ¡������ | �ֹμ�                            
            pGDColumn[137] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("MINUS_SP_TAX_AMOUNT");       // ������õ¡������ | ��Ư��                            
            pGDColumn[138] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TOTAL_MINUS_TAX");           // ������õ¡������ �հ�


            pGDColumn[139] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRINT_DATE");                // ��³�¥
            pGDColumn[140] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRINT_OWNER");               // ¡���ǹ���
            pGDColumn[141] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("OFFICER_FLAG");              // �ӿ�����
            pGDColumn[142] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_1");           // ��������
            pGDColumn[143] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_2");           // �����ذ�
            pGDColumn[144] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_3");           // �ڹ�������
            pGDColumn[145] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_4");           // �ӿ�����
            pGDColumn[146] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_5");           // �߰�����
            pGDColumn[147] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIRE_REASON_6");           // ��Ÿ
            pGDColumn[148] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_MONTH1");                // ���������޿� �������

            pGDColumn[149] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("ADD_MONTH2");                // ���� �� �����޿� �������.
            pGDColumn[150] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RETIREMENT_PENSION_ACCOUNT");// �������� ���¹�ȣ.
            pGDColumn[151] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("DEFERRAL_TAX_AMT");          // �����̿� �ݾ�.
            pGDColumn[152] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("PRE_RECEIPT_RETIRE_PAY_AMT");// ������� �����޿���.
            pGDColumn[153] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RP_CORP_DESC");              // �������� ����ڸ�.
            pGDColumn[154] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RP_TAX_REG_NUM");            // ����ڹ�ȣ
            pGDColumn[155] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("RP_ACCOUNT_NUM");            // ���¹�ȣ
            pGDColumn[156] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TRANS_RETIRE_AMOUNT");       // �Ա�(��ü) ���������޿�
            pGDColumn[157] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TRANS_HONORARY_AMOUNT");     // �Ա�(��ü) ���������޿��̿�
            pGDColumn[158] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TRANS_DATE");                // �Ա�(��ü) ��

            pGDColumn[159] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("EXPIRE_DATE");               // ������
            pGDColumn[160] = pGrid_WITHHOLDING_TAX.GetColumnToIndex("TRANS_INCOME_TAX_AMOUNT");   // �����̿� �������� �հ�


            //---------------------------------------------------------------------------------------------------------------------

            pXLColumn[0]   = 37;   // ���� ����(������1)    
            pXLColumn[1]   = 42;   // ���� ����(������2)      
            pXLColumn[2]   = 37;   // ���ܱ��� ����(������1)                                    
            pXLColumn[3]   = 42;   // ���ܱ��� ����(�ܱ���9)                              
            pXLColumn[4]   = 31;   // ��������                                          
            pXLColumn[5]   = 39;   // ���������ڵ�                                          
            pXLColumn[6]   = 33;   // ¡���ǹ��ڱ���                                                  
            pXLColumn[7]   = 11;   // ����ڵ�Ϲ�ȣ                                              
            pXLColumn[8]   = 27;   // ���θ�(��ȣ)                                                   
            pXLColumn[9]   = 39;   // ��ǥ��(����)                                  
            pXLColumn[10]  = 27;   // ������(�ּ�)   

            pXLColumn[11]  = 11;   // ����                      
            pXLColumn[12]  = 27;   // �ֹι�ȣ                                     
            pXLColumn[13]  = 11;   // �ּ�                                     
            pXLColumn[14]  = 11;   // �ͼӿ��� ���� ����                                    
            pXLColumn[15]  = 27;   // ��������                                  �Ⱦ�                        
            pXLColumn[16]  = 11;   // �ͼӿ��� ������ ����(��������)                                              
            pXLColumn[17]  = 20;   // �������װ�������1                         �Ⱦ�                      
            pXLColumn[18]  = 28;   // �������װ�������2                         �Ⱦ�                
            pXLColumn[19]  = 36;   // �������װ�������3                         �Ⱦ�               
            pXLColumn[20]  = 12;   // �������װ������� �հ�                     �Ⱦ�

            pXLColumn[21]  = 12;   // �ٹ�ó��1                                       
            pXLColumn[22]  = 20;   // �ٹ�ó��2                                       
            pXLColumn[23]  = 28;   // �ٹ�ó��3                                   
            pXLColumn[24]  = 36;   // �ٹ�ó�� �հ�                             �Ⱦ�                                         
            pXLColumn[25]  = 12;   // ����ڵ�Ϲ�ȣ1                                             
            pXLColumn[26]  = 20;   // ����ڵ�Ϲ�ȣ2                                             
            pXLColumn[27]  = 28;   // ����ڵ�Ϲ�ȣ3                                          
            pXLColumn[28]  = 36;   // ����ڵ�Ϲ�ȣ �հ�                       �Ⱦ�                          
            pXLColumn[29]  = 12;   // �����޿�1                       
            pXLColumn[30]  = 20;   // �����޿�2 

            pXLColumn[31] = 28;   // �����޿�3                        
            pXLColumn[32] = 36;   // �����޿� �հ�                                      
            pXLColumn[33] = 16;   // ����������(�߰�������)1                                             
            pXLColumn[34] = 24;   // ����������(�߰�������)2 
            pXLColumn[35] = 32;   // ����������(�߰�������)3
            pXLColumn[36] = 36;   // ����������(�߰�������) �հ�       
            pXLColumn[37] = 12;   // ���������Ͻñ�1                                                   
            pXLColumn[38] = 20;   // ���������Ͻñ�2                                                   
            pXLColumn[39] = 28;   // ���������Ͻñ�3                                               
            pXLColumn[40] = 36;   // ���������Ͻñ� �հ�    

            pXLColumn[41]  = 16;   // ���������� ���������Ͻñ�                                      
            pXLColumn[42]  = 24;   // ���������� ���������Ͻñ�                                 
            pXLColumn[43]  = 32;   // ���������� ���������Ͻñ�                           
            pXLColumn[44]  = 36;   // ���������� ���������Ͻñ� �հ�  
            pXLColumn[45]  = 12;   // ������1                                        
            pXLColumn[46]  = 20;   // ������2                                    
            pXLColumn[47]  = 28;   // ������3                              
            pXLColumn[48]  = 36;   // ������ �հ�                                        
            pXLColumn[49]  = 16;   // �����ܰ�1                                             
            pXLColumn[50]  = 24;   // �����ܰ�2

            pXLColumn[51]  = 32;   // �����ܰ�3                                        
            pXLColumn[52]  = 36;   // �����ܰ� �հ�                             
            pXLColumn[53]  = 12;   // ������ҵ�1                                        
            pXLColumn[54]  = 20;   // ������ҵ�2                  
            pXLColumn[55]  = 28;   // ������ҵ�3                                              
            pXLColumn[56]  = 36;   // ������ҵ� �հ�                                    
            pXLColumn[57]  = 17;   // �Ѽ��ɾ�1                                      
            pXLColumn[58]  = 22;   // ������ �հ��1                                 
            pXLColumn[59]  = 27;   // �ҵ��� ���Ծ�1                             
            pXLColumn[60]  = 32;   // �������� �ҵ������1        

            pXLColumn[61]  = 37;   // ���������Ͻñ�1                   
            pXLColumn[62]  = 17;   // �Ѽ��ɾ�2                   
            pXLColumn[63]  = 22;   // ������ �հ��2                                              
            pXLColumn[64]  = 27;   // �ҵ��� ���Ծ�2                                    
            pXLColumn[65]  = 32;   // �������� �ҵ������2                                       
            pXLColumn[66]  = 37;   // ���������Ͻñ�2                                   
            pXLColumn[67]  = 9;   // ���������Ͻñ� ���޿����, �̿��ݾ�                                 
            pXLColumn[68]  = 21;   // ���Ͻñ�                               
            pXLColumn[69]  = 25;   // ���ɰ��������޿���               
            pXLColumn[70]  = 29;   // ȯ�������ҵ����         

            pXLColumn[71] = 33;   // ȯ�������ҵ����ǥ��                                               
            pXLColumn[72] = 37;   // ȯ�꿬��� ����ǥ��                                             
            pXLColumn[73] = 41;   // ȯ�� ����� ���⼼��                                            
            pXLColumn[74] = 9;   // ���������Ͻñ� ���޿����, �̿��ݾ�                                        
            pXLColumn[75] = 9;   // ���������Ͻñ� ���޿����, �̿��ݾ�                                             
            pXLColumn[76] = 21;   // ���Ͻñ�                                               
            pXLColumn[77] = 25;   // ���ɰ��������޿���                                             
            pXLColumn[78] = 29;   // ȯ�������ҵ����                                             
            pXLColumn[79] = 33;   // ȯ�������ҵ����ǥ��                                          
            pXLColumn[80] = 37;   // ȯ�꿬��� ����ǥ��
   
            pXLColumn[81] = 41;   // ȯ�� ����� ���⼼��                                               
            pXLColumn[82] = 9;   // ���������Ͻñ� ���޿����, �̿��ݾ�                                             
            pXLColumn[83] = 12;   // �Ի���(���������)1  
            pXLColumn[84] = 15;   // �����1                                               
            pXLColumn[85] = 18;   // �ټӿ���1                                                
            pXLColumn[86] = 20;   // ���ܿ���1                                             
            pXLColumn[87] = 26;   // �ټӿ���1                                             
            pXLColumn[88] = 28;   // �Ի���2
            pXLColumn[89] = 31;   // �����2            
            pXLColumn[90] = 34;   // �ټӿ���2 

            pXLColumn[91]  = 36;   // ���ܿ���2                                            
            pXLColumn[92]  = 42;   // �ټӿ���2
            pXLColumn[93]  = 9;   // �Ի���(���������)1                                    
            pXLColumn[94]  = 12;   // �����1                                     
            pXLColumn[95]  = 15;   // �ټӿ���1                                      
            pXLColumn[96]  = 18;   // ���ܿ���1
            pXLColumn[97]  = 28;   // �Ի���2                 
            pXLColumn[98]  = 31;   // �����2                          
            pXLColumn[99]  = 34;   // �ټӿ���2        
            pXLColumn[100] = 36;   // ���ܿ���2  

            pXLColumn[101] = 24;   // �ߺ�����1   ??                        
            pXLColumn[102] = 24;   // �ߺ�����2   ??      
            pXLColumn[103] = 12;   // �����޿���     
            pXLColumn[104] = 25;   // ��������  ??                                
            pXLColumn[105] = 38;   // �����޿��� �հ�               
            pXLColumn[106] = 12;   // �����ҵ���� - ��           
            pXLColumn[107] = 25;   // �����ҵ���� - ��                                 
            pXLColumn[108] = 38;   // �����ҵ���� �հ�   
            pXLColumn[109] = 12;   // �����ҵ����ǥ�� - ���������޿�  
            pXLColumn[110] = 25;   // �����ҵ����ǥ�� - �����̿� �����޿�

            pXLColumn[111] = 38;   // �����ҵ����ǥ�� �հ�                
            pXLColumn[112] = 12;   // ���װ��ٰ� - ����հ���ǥ�� - ���������޿�              
            pXLColumn[113] = 25;   // ���װ��ٰ� - ����հ���ǥ�� - �����̿� �����޿�                                        
            pXLColumn[114] = 38;   // ����հ���ǥ�� �հ�.                        
            pXLColumn[115] = 12;   // ���װ��ٰ� - ����ջ��⼼�� - ���������޿�                     
            pXLColumn[116] = 25;   // ���װ��ٰ� - ����ջ��⼼�� - �����̿� �����޿�                                
            pXLColumn[117] = 38;   // ����ջ��⼼�� �հ�                                       
            pXLColumn[118] = 12;   // ���װ��ٰ� - ���⼼�� - ���������޿�                        
            pXLColumn[119] = 25;   // ���װ��ٰ� - ���⼼�� - �����̿� �����޿�                     
            pXLColumn[120] = 28;   // ���⼼�� �հ�                  

            pXLColumn[121] = 12;   // ���װ��ٰ� - ���װ���(�ܱ�����) - ���������޿�                  
            pXLColumn[122] = 25;   // ���װ��ٰ� - ���װ���(�ܱ�����) - �����̿� �����޿�                  
            pXLColumn[123] = 28;   // ���װ��� �հ�                             
            //pXLColumn[124] = 27;   // �������� - �����ҵ漼 - ���������޿�                            
            //pXLColumn[125] = 35;   // �������� - �����ҵ漼 - �����̿� �����޿� 
            //pXLColumn[126] = 29;   // �������� �հ� 
            pXLColumn[127] = 12;   // �������� - �ҵ漼 �հ�
            pXLColumn[128] = 20;   // �������� - �ֹμ� �հ�    
            pXLColumn[129] = 27;   // ����� Ư���� �հ�   
            pXLColumn[130] = 35;   // �������� �հ�  

            pXLColumn[131] = 12;   // ��(��)�ٹ��� �ⳳ�μ��� | �ҵ漼               
            pXLColumn[132] = 20;   // ��(��)�ٹ��� �ⳳ�μ��� | �ֹμ�              
            pXLColumn[133] = 27;   // ��(��)�ٹ��� �ⳳ�μ��� | ��Ư��                                       
            pXLColumn[134] = 35;   // ��(��)�ٹ��� �ⳳ�μ��� | �հ�                         
            pXLColumn[135] = 12;   // ������õ¡������ | �ҵ漼                     
            pXLColumn[136] = 20;   // ������õ¡������ | �ֹμ�                                  
            pXLColumn[137] = 27;   // ������õ¡������ | ��Ư��                                    
            pXLColumn[138] = 35;   // ������õ¡������ �հ�                  
            pXLColumn[139] = 35;   // ��³�¥                    
            pXLColumn[140] = 27;   // ¡���ǹ���

            pXLColumn[141] = 39;   // �ӿ�����              
            pXLColumn[142] = 28;   // ��������          
            pXLColumn[143] = 34;   // �����ذ�                                    
            pXLColumn[144] = 40;   // �ڹ�������                         
            pXLColumn[145] = 28;   // �ӿ�����                     
            pXLColumn[146] = 34;   // �߰�����                                  
            pXLColumn[147] = 40;   // ��Ÿ
            pXLColumn[148] = 22;   // ���������޿� �������           
            pXLColumn[149] = 38;   // ���� �� �����޿� �������.  
            pXLColumn[150] = 12;   // �������� ���¹�ȣ.

            pXLColumn[151] = 13;   // �����̿� �ݾ�.              
            pXLColumn[152] = 17;   // ������� �����޿���.      
            pXLColumn[153] = 3;   // �������� ����ڸ�.                         
            pXLColumn[154] = 9;   // ����ڹ�ȣ                  
            pXLColumn[155] = 15;   // ���¹�ȣ          
            pXLColumn[156] = 21;   // �Ա�(��ü) ���������޿�                              
            pXLColumn[157] = 26;   // �Ա�(��ü) ���������޿��̿�
            pXLColumn[158] = 31;   // �Ա�(��ü) ��    
            pXLColumn[159] = 39;   // ������
            pXLColumn[160] = 39;   // �����̿� �������� �հ�

        }

        #endregion;

        #region ----- SetArray2 ----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_2013, out int[] pGDColumn, out int[] pXLColumn)
        {
            pGDColumn = new int[135];
            pXLColumn = new int[135];

            //--------------------------------------------------------------------------------------------------------------------

            // ���� ����� �׸�(���ֱ���, ��/�ܱ���, ��������)
            pGDColumn[0] = pGrid_PRINT_2013.GetColumnToIndex("RESIDENT_TYPE_1");                           // ���� ����(������1)       
            pGDColumn[1] = pGrid_PRINT_2013.GetColumnToIndex("RESIDENT_TYPE_2");                           // ���� ����(������2)       
            pGDColumn[2] = pGrid_PRINT_2013.GetColumnToIndex("NATIONALITY_TYPE_1");                        // ���ܱ��� ����(������1)   
            pGDColumn[3] = pGrid_PRINT_2013.GetColumnToIndex("NATIONALITY_TYPE_9");                        // ���ܱ��� ����(�ܱ���9)      
            pGDColumn[4] = pGrid_PRINT_2013.GetColumnToIndex("NATION_DESC");                               // ��������
            pGDColumn[5] = pGrid_PRINT_2013.GetColumnToIndex("NATION_ISO_CODE");                           // ���������ڵ�
            pGDColumn[6] = pGrid_PRINT_2013.GetColumnToIndex("OWNER_TYPE");                                // ¡���ǹ��ڱ���
            
            // ¡���ǹ��� ��������
            pGDColumn[7] = pGrid_PRINT_2013.GetColumnToIndex("VAT_NUMBER");                                // ����ڵ�Ϲ�ȣ                                       
            pGDColumn[8] = pGrid_PRINT_2013.GetColumnToIndex("CORP_NAME");                                 // ���θ�(��ȣ)                                         
            pGDColumn[9] = pGrid_PRINT_2013.GetColumnToIndex("PRESIDENT_NAME");                            // ��ǥ��(����)   
            pGDColumn[10] = pGrid_PRINT_2013.GetColumnToIndex("LEGAL_NUMBER");                             // ���ι�ȣ 
            pGDColumn[11] = pGrid_PRINT_2013.GetColumnToIndex("ORG_ADDRESS");                              // ������(�ּ�)   

            //�ҵ��� ��������
            pGDColumn[12] = pGrid_PRINT_2013.GetColumnToIndex("NAME");                                     // ����                                                 
            pGDColumn[13] = pGrid_PRINT_2013.GetColumnToIndex("REPRE_NUM");                                // �ֹι�ȣ                                             
            pGDColumn[14] = pGrid_PRINT_2013.GetColumnToIndex("ADDRESS");                                  // �ּ�        
            pGDColumn[15] = pGrid_PRINT_2013.GetColumnToIndex("OFFICER_FLAG");                             // �ӿ�����
            pGDColumn[16] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_PENSION_REGIST_DATE");               // Ȯ���޿��� �������� ���� ������.    
            pGDColumn[17] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE");                              // ������ ��¥.                 
            pGDColumn[18] = pGrid_PRINT_2013.GetColumnToIndex("START_RETIRE_DATE");                        // �ͼӿ��� ���� ����                               
            pGDColumn[19] = pGrid_PRINT_2013.GetColumnToIndex("LAST_RETIRE_DATE");                         // �ͼӿ��� ������ ����(��������).                                 
            pGDColumn[20] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_1");                          // ��������                                 
            pGDColumn[21] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_2");                          // �����ذ�  
            pGDColumn[22] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_3");                          // �ڹ�������      
            pGDColumn[23] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_4");                          // �ӿ�����                                            
            pGDColumn[24] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_5");                          // �߰�����                                            
            pGDColumn[25] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_REASON_6");                          // ��Ÿ    

            //�����޿���Ȳ.
            pGDColumn[26] = pGrid_PRINT_2013.GetColumnToIndex("WORK_CORP_NAME1");                         // �ٹ�ó�� ( �߰�����)                  
            pGDColumn[27] = pGrid_PRINT_2013.GetColumnToIndex("WORK_CORP_NAME2");                         // �ٹ�ó�� (������)                                   
            pGDColumn[28] = pGrid_PRINT_2013.GetColumnToIndex("WORK_VAT_NUMBER1");                         // ����ڵ�Ϲ�ȣ ( �߰�����)                                    
            pGDColumn[29] = pGrid_PRINT_2013.GetColumnToIndex("WORK_VAT_NUMBER2");                         // ����ڵ�Ϲ�ȣ (������)                             
            pGDColumn[30] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_TOTAL_AMOUNT1");                     // �����޿� (�߰�����)
            pGDColumn[31] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_TOTAL_AMOUNT2");                     // �����޿� (������)         
            pGDColumn[32] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_TOTAL_AMOUNT3");                     // �����޿� (����)
            pGDColumn[33] = pGrid_PRINT_2013.GetColumnToIndex("NON_TAX_TOTAL_AMOUNT1");                    // ����� �����޿� (�߰�����)             
            pGDColumn[34] = pGrid_PRINT_2013.GetColumnToIndex("NON_TAX_TOTAL_AMOUNT2");                    // ����� �����޿� (������)                        
            pGDColumn[35] = pGrid_PRINT_2013.GetColumnToIndex("NON_TAX_TOTAL_AMOUNT3");                    // ���� �����޿� (����)
            pGDColumn[36] = pGrid_PRINT_2013.GetColumnToIndex("TAX_RETIRE_TOTAL_AMOUNT1");                 // ������� �����޿� (�߰�����)
            pGDColumn[37] = pGrid_PRINT_2013.GetColumnToIndex("TAX_RETIRE_TOTAL_AMOUNT2");                 // ������� ������ (������)                       
            pGDColumn[38] = pGrid_PRINT_2013.GetColumnToIndex("TAX_RETIRE_TOTAL_AMOUNT3");                 // ������� �����޿� (����)      

            //�ټӿ���.
            pGDColumn[39] = pGrid_PRINT_2013.GetColumnToIndex("ORI_JOIN_DATE1");                           //�Ի��� ( �߰����� �ټӿ���)                              
            pGDColumn[40] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR1");                          // �����/��������� ( �߰����� �ټӿ���)                                      
            pGDColumn[41] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE1");                             // ����� ( �߰����� �ټӿ���)
            pGDColumn[42] = pGrid_PRINT_2013.GetColumnToIndex("CLOSED_DATE1");                             // ������ ( �߰����� �ټӿ���)
            pGDColumn[43] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH1");                              // �ټӿ��� ( �߰����� �ټӿ���)
            pGDColumn[44] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH1");                            // ���ܿ��� ( �߰����� �ټӿ���)
            pGDColumn[45] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH1");                            // �������� ( �߰����� �ټӿ���)                                    
            pGDColumn[46] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR1");                               // �ټӿ��� ( �߰����� �ټӿ���)


            pGDColumn[47] = pGrid_PRINT_2013.GetColumnToIndex("ORI_JOIN_DATE2");                           // �Ի���(������)                                               
            pGDColumn[48] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR2");                          // �����/��������� ( ������)
            pGDColumn[49] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE2");                             // ����� ( ������ )                                              
            pGDColumn[50] = pGrid_PRINT_2013.GetColumnToIndex("CLOSED_DATE2");                             // ������ ( ������ )
            pGDColumn[51] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH2");                              // �ټӿ��� ( ������)                                                   
            pGDColumn[52] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH2");                            // �����ܰ� �հ� 
            pGDColumn[53] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH2");                            // ������ҵ�1     
            pGDColumn[54] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR2");                               // ������ҵ�2  


            pGDColumn[55] = pGrid_PRINT_2013.GetColumnToIndex("ORI_JOIN_DATE3");                           // �Ի���(����(�ջ�))                                   
            pGDColumn[56] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR3");                          // �����/��������� (����(�ջ�))                                   
            pGDColumn[57] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE3");                             // ����� (����(�ջ�))
            pGDColumn[58] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH3");                              // �ټӿ��� (����(�ջ�))
            pGDColumn[59] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH3");                            // ���ܿ��� (����(�ջ�))                                    
            pGDColumn[60] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH3");                            // �������� (����(�ջ�))                             
            pGDColumn[61] = pGrid_PRINT_2013.GetColumnToIndex("PRE_LONG_YEAR3");                           // �ߺ����� (����(�ջ�))                                  
            pGDColumn[62] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR3");                               // �ټӿ��� (����(�ջ�)) 

            pGDColumn[63] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR4");                          // ����� (2012.12.31 ����)                                  
            pGDColumn[64] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE4");                             // �����  (2012.12.31 ����)                            
            pGDColumn[65] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH4");                              // �ټӿ��� (2012.12.31 ����)     
            pGDColumn[66] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH4");                            // ���ܿ��� (2012.12.31 ����)
            pGDColumn[67] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH4");                            // �������� (2012.12.31 ����)           
            pGDColumn[68] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR4");                               // �ټӳ�� (2012.12.31 ����)

            pGDColumn[69] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE_FR5");                          // ����� ( 2013.01.01 ����)                                   
            pGDColumn[70] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_DATE5");                             // �����  ( 2013.01.01 ����)
            pGDColumn[71] = pGrid_PRINT_2013.GetColumnToIndex("LONG_MONTH5");                              // �ټӿ��� ( 2013.01.01 ����)                         
            pGDColumn[72] = pGrid_PRINT_2013.GetColumnToIndex("EXCEPT_MONTH5");                            // ���ܿ��� ( 2013.01.01 ����)   
            pGDColumn[73] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_MONTH5");                            // �������� ( 2013.01.01 ����)
            pGDColumn[74] = pGrid_PRINT_2013.GetColumnToIndex("LONG_YEAR5");                               // �ټӳ�� ( 2013.01.01 ����)

            //�����ҵ����ǥ�ذ��
            pGDColumn[75] = pGrid_PRINT_2013.GetColumnToIndex("MID_RETIRE_AMOUNT");                        // �����ҵ�(�߰�����)   
            pGDColumn[76] = pGrid_PRINT_2013.GetColumnToIndex("RETIRE_AMOUNT");                            // �����ҵ�(������)   
            pGDColumn[77] = pGrid_PRINT_2013.GetColumnToIndex("SUM_RETIRE_AMOUNT");                        // �����ҵ�(����(�ջ�)        
            pGDColumn[78] = pGrid_PRINT_2013.GetColumnToIndex("INCOME_DED_AMOUNT");                        // �����ҵ���������     
            pGDColumn[79] = pGrid_PRINT_2013.GetColumnToIndex("LONG_DED_AMOUNT");                          // �ټӿ�������                           
            pGDColumn[80] = pGrid_PRINT_2013.GetColumnToIndex("TAX_STD_AMOUNT");                           // �����ҵ����ǥ��(27-28-29)    


            //�����ҵ漼�װ��
            pGDColumn[81] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_ANBON1");                           // ����ǥ�ؾȺ�(2012.12.31����)          
            pGDColumn[82] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_STD1");                             // ����հ���ǥ��(2012.12.31����)
            pGDColumn[83] = pGrid_PRINT_2013.GetColumnToIndex("AVG_YEAR_COMP_TAX1");                       // ����ջ��⼼��(2012.12.31����)                          
            pGDColumn[84] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX1");                                // ���⼼�� (2012.12.31����)

            pGDColumn[85] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_ANBON2");                           // ����ǥ�ؾȺ�(2013.01.01���� )                                      
            pGDColumn[86] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_STD2");                             // ����ǥ��(2013.01.01���� )                                          
            pGDColumn[87] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_AVG_TAX_2");                         // ȯ�����ǥ�� (2013.01.01���� )                          
            pGDColumn[88] = pGrid_PRINT_2013.GetColumnToIndex("CAHNGE_AVG_COMP_TAX2");                     // ȯ����⼼�� (2013.01.01���� )                       
            pGDColumn[89] = pGrid_PRINT_2013.GetColumnToIndex("AVG_YEAR_COMP_TAX2");                       // ����ջ��⼼�� (2013.01.01���� )                                         
            pGDColumn[90] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX2");                                // ���⼼�� (2013.01.01���� )

            pGDColumn[91] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_ANBON_SUM");                        // ����ǥ�ؾȺ�(����(�ջ�))                                
            pGDColumn[92] = pGrid_PRINT_2013.GetColumnToIndex("AVG_TAX_STD_SUM");                          // ����հ���ǥ��(����(�ջ�))                                    
            pGDColumn[93] = pGrid_PRINT_2013.GetColumnToIndex("CHANGE_AVG_TAX_SUM");                       // ȯ�����ǥ�� 
            pGDColumn[94] = pGrid_PRINT_2013.GetColumnToIndex("CAHNGE_AVG_COMP_TAX_SUM");                  // ȯ����⼼�׾�                                           
            pGDColumn[95] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX_AMOUNT_SUM");                      // ����ջ��⼼��                                           
            pGDColumn[96] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX_SUM");                             // ���⼼��                                     
            pGDColumn[97] = pGrid_PRINT_2013.GetColumnToIndex("AHEAD_TAX");                                // �ⳳ�μ���                                
            pGDColumn[98] = pGrid_PRINT_2013.GetColumnToIndex("REPORT_TAX_SUM");                           // �Ű��󼼾� 

            //�̿������ҵ漼�װ��
            pGDColumn[99] = pGrid_PRINT_2013.GetColumnToIndex("REPORT_TAX_SUM_39");                        // �Ű��󼼾� (39)                                  
            pGDColumn[100] = pGrid_PRINT_2013.GetColumnToIndex("RETIREMENT_CORP_NAME");                    // ���ݰ��������
            pGDColumn[101] = pGrid_PRINT_2013.GetColumnToIndex("TAX_REG_NUM");                             // ����ڵ�Ϲ�ȣ        
            pGDColumn[102] = pGrid_PRINT_2013.GetColumnToIndex("ACCOUNT_NUM");                             // ���¹�ȣ
            pGDColumn[103] = pGrid_PRINT_2013.GetColumnToIndex("ISSUE_DATE");                              // �Ա���                                           
            pGDColumn[104] = pGrid_PRINT_2013.GetColumnToIndex("TRANS_ACCOUNT_AMOUNT_1");                  // �����Աݱݾ�                                    
            pGDColumn[105] = pGrid_PRINT_2013.GetColumnToIndex("RETIREMENT_PAY");                          // �����޿�                                
            pGDColumn[106] = pGrid_PRINT_2013.GetColumnToIndex("TRANS_INCOME_TAX_AMOUNT");                 // �̿������ҵ漼        

            //���θ�
            pGDColumn[107] = pGrid_PRINT_2013.GetColumnToIndex("INCOME_TAX_AMOUNT");                       // �Ű��󼼾�(�ҵ漼)                            
            pGDColumn[108] = pGrid_PRINT_2013.GetColumnToIndex("RESIDENT_TAX_AMOUNT");                     // �Ű��󼼾�(����ҵ漼)
            pGDColumn[109] = pGrid_PRINT_2013.GetColumnToIndex("SP_TAX_AMOUNT");                           // �Ű��󼼾�(�����Ư����)
            pGDColumn[110] = pGrid_PRINT_2013.GetColumnToIndex("SUM_INCOME_RESIDENT");                     // �Ű��󼼾�(��)

            pGDColumn[111] = pGrid_PRINT_2013.GetColumnToIndex("DEFERRED_INCOME_TAX");                     // �̿������ҵ漼(�ҵ漼)           
            pGDColumn[112] = pGrid_PRINT_2013.GetColumnToIndex("DEFERRED_RESIDENT_TAX");                   // �̿������ҵ漼(����ҵ漼)    
            pGDColumn[113] = pGrid_PRINT_2013.GetColumnToIndex("DEFERRED_SP_TAX");                         // �̿������ҵ漼(�����Ư����)    
            pGDColumn[114] = pGrid_PRINT_2013.GetColumnToIndex("DEFERRED_SUM_INCOME_RESIDENT");            // �̿������ҵ漼(��)     

            pGDColumn[115] = pGrid_PRINT_2013.GetColumnToIndex("DEDUCTION_INCOME_TAX");                    // ������û¡������(�ҵ漼)
            pGDColumn[116] = pGrid_PRINT_2013.GetColumnToIndex("DEDUCTION_RESIDENT_TAX");                  // ������û¡������(����ҵ漼)  
            pGDColumn[117] = pGrid_PRINT_2013.GetColumnToIndex("DEDUCTION_SP_TAX");                        // ������û¡������(�����Ư����)  
            pGDColumn[118] = pGrid_PRINT_2013.GetColumnToIndex("DEDUCTION_SUM_INCOME_RESIDENT");           // ������û¡������(��)  

            //�ϴ�����Ʈ.
            pGDColumn[119] = pGrid_PRINT_2013.GetColumnToIndex("PRINT_DATE");                              // ������� 
            pGDColumn[120] = pGrid_PRINT_2013.GetColumnToIndex("PRINT_OWNER");                             // ��ǥ�̻�                               
            pGDColumn[121] = pGrid_PRINT_2013.GetColumnToIndex("FINAL_RETIRE_DATE");                       // ��������

            //2016 �߰���
            pGDColumn[122] = pGrid_PRINT_2013.GetColumnToIndex("SUM_RETIRE_AMOUNT_3");                     // 31.�����ҵ�(����(�ջ�))
            pGDColumn[123] = pGrid_PRINT_2013.GetColumnToIndex("LONG_DED_AMOUNT_3");                       // 32.�ټӿ�������
            pGDColumn[124] = pGrid_PRINT_2013.GetColumnToIndex("CHG_STD_AMOUNT_3");                        // 33.ȯ��޿�
            pGDColumn[125] = pGrid_PRINT_2013.GetColumnToIndex("CHG_DED_AMOUNT_3");                        // 34.ȯ��޿�������
            pGDColumn[126] = pGrid_PRINT_2013.GetColumnToIndex("CHG_TAX_STD_AMOUNT_3");                    // 35.�����ҵ����ǥ��
            pGDColumn[127] = pGrid_PRINT_2013.GetColumnToIndex("CHG_COMP_TAX_AMOUNT_3");                   // 42.ȯ����⼼��
            pGDColumn[128] = pGrid_PRINT_2013.GetColumnToIndex("COMP_TAX_AMOUNT_3");                       // 43.���⼼��
            pGDColumn[129] = pGrid_PRINT_2013.GetColumnToIndex("ADJUSTMENT_YYYY_3");                       // 44.�������� ���ϴ� ��������
            pGDColumn[130] = pGrid_PRINT_2013.GetColumnToIndex("REAL_COMP_TAX_AMOUNT_3");                  // 45.Ư��������⼼��
            pGDColumn[131] = pGrid_PRINT_2013.GetColumnToIndex("PREPAID_TAX_AMOUNT_3");                    // 46.�ⳳ�� ����
            pGDColumn[132] = pGrid_PRINT_2013.GetColumnToIndex("INCOME_TAX_AMOUNT_3");                     // 47.�Ű��󼼾�
            pGDColumn[133] = pGrid_PRINT_2013.GetColumnToIndex("TRANS_ACCOUNT_SUM_AMOUNT");                // ���ݰ��� �Ա��հ� 
            pGDColumn[134] = pGrid_PRINT_2013.GetColumnToIndex("TAX_OFFICE_NAME");                         // ��������
            

            //---------------------------------------------------------------------------------------------------------------------

            pXLColumn[0] = 37;   // ���� ����(������1)    
            pXLColumn[1] = 42;   // ���� ����(������2)   
   
            pXLColumn[2] = 37;   // ���ܱ��� ����(������1)                                    
            pXLColumn[3] = 42;   // ���ܱ��� ����(�ܱ���9)   
                           
            pXLColumn[4] = 31;   // ��������                                          
            pXLColumn[5] = 39;   // ���������ڵ�          
                                
            pXLColumn[6] = 33;   // ¡���ǹ��ڱ���  
            
            //---------------------------------------
            pXLColumn[7] = 11;   // ����ڵ�Ϲ�ȣ                                              
            pXLColumn[8] = 27;   // ���θ�(��ȣ)                                                   
            pXLColumn[9] = 39;   // ��ǥ��(����)

            pXLColumn[10] = 11;   // ���ι�ȣ 
            pXLColumn[11] = 27;   // ������(�ּ�)   

            //-----------------------------------------
            pXLColumn[12] = 11;   // ����                      
            pXLColumn[13] = 27;   // �ֹι�ȣ     
                                
            pXLColumn[14] = 11;   // �ּ�                                     
            pXLColumn[15] = 39;   // �ӿ�����

            pXLColumn[16] = 11;   // Ȯ���޿��� �������� ���� ������
            pXLColumn[17] = 36;   // ������ ��¥                    
                         
            pXLColumn[18] = 11;   // �ͼӿ��� ���� ����            
            pXLColumn[19] = 11;   // �ͼӿ��� ������ ����(��������).                
            pXLColumn[20] = 26;   // ��������                
            pXLColumn[21] = 32;   // �����ذ�  
            pXLColumn[22] = 38;   // �ڹ�������                                        
            pXLColumn[23] = 26;   // �ӿ�����                                      
            pXLColumn[24] = 32;   // �߰�����                                    
            pXLColumn[25] = 38;   // ��Ÿ
            //-----------------------------------------

            pXLColumn[26] = 14;   // �ٹ�ó�� ( �߰�����)                                             
            pXLColumn[27] = 24;   // �ٹ�ó�� ( �߰�����)         
                                    
            pXLColumn[28] = 14;   // ����ڵ�Ϲ�ȣ ( �߰�����)                                            
            pXLColumn[29] = 24;   // ����ڵ�Ϲ�ȣ (������)  
           
            pXLColumn[30] = 14;   // �����޿� (�߰�����)                   
            pXLColumn[31] = 24;   // �����޿� (������)                        
            pXLColumn[32] = 34;   // �����޿� (����)    

            pXLColumn[33] = 14;   // ����� �����޿� (�߰�����)                                            
            pXLColumn[34] = 24;   // ����� �����޿� (������)      
            pXLColumn[35] = 34;   // ���� �����޿� (����)

            pXLColumn[36] = 14;   // ������� �����޿� (�߰�����) 
            pXLColumn[37] = 24;   // ������� ������ (������)                                                   
            pXLColumn[38] = 34;   // ������� �����޿� (����)      

            //-----------------------------------------

            pXLColumn[39] = 13;   // �Ի��� ( �߰����� �ټӿ���)                                             
            pXLColumn[40] = 17;   // �����/��������� ( �߰����� �ټӿ���)    
            pXLColumn[41] = 21;   // ����� ( �߰����� �ټӿ���)                                
            pXLColumn[42] = 25;   // ������ ( �߰����� �ټӿ���)                            
            pXLColumn[43] = 29;   // �ټӿ��� ( �߰����� �ټӿ���)                       
            pXLColumn[44] = 32;   // ���ܿ��� ( �߰����� �ټӿ���)  
            pXLColumn[45] = 35;   // ������� ( �߰����� �ټӿ���)                                        
            pXLColumn[46] = 41;   // �ټӿ��� ( �߰����� �ټӿ���)  

            pXLColumn[47] = 13;   // �Ի���( ������)                               
            pXLColumn[48] = 17;   // �����/��������� ( ������)                              
            pXLColumn[49] = 21;   // ����� ( ������ )                                                
            pXLColumn[50] = 25;   // ������ ( ������ )
            pXLColumn[51] = 29;   // �ټӿ��� ( ������)                                            
            pXLColumn[52] = 32;   // ���ܿ��� ( ������)                            
            pXLColumn[53] = 35;   // ������� ( ������)                                       
            pXLColumn[54] = 41;   // �ټӿ��� ( ������)  

            pXLColumn[55] = 13;   // �Ի���(����(�ջ�))                                              
            pXLColumn[56] = 17;   // �����/��������� (����(�ջ�))                                   
            pXLColumn[57] = 21;   // ����� (����(�ջ�))                                 
            pXLColumn[58] = 29;   // �ټӿ��� (����(�ջ�))                             
            pXLColumn[59] = 32;   // ���ܿ��� (����(�ջ�))                                
            pXLColumn[60] = 35;   // ������� (����(�ջ�))       
            pXLColumn[61] = 38;   // �ߺ����� (����(�ջ�))            
            pXLColumn[62] = 41;   // �ټӿ��� (����(�ջ�))    

            pXLColumn[63] = 17;   // ����� (2012.12.31 ����)                                      
            pXLColumn[64] = 21;   // �����  (2012.12.31 ����)                              
            pXLColumn[65] = 29;   // �ټӿ��� (2012.12.31 ����)                                  
            pXLColumn[66] = 32;   // ���ܿ��� (2012.12.31 ����)                             
            pXLColumn[67] = 35;   // ������� (2012.12.31 ����)                                  
            pXLColumn[68] = 41;   // �ټӳ�� (2012.12.31 ����)

            pXLColumn[69] = 17;   // ����� ( 2013.01.01 ����)                   
            pXLColumn[70] = 21;   // �����  ( 2013.01.01 ����)         
            pXLColumn[71] = 29;   // �ټӿ��� ( 2013.01.01 ����)                                             
            pXLColumn[72] = 32;   // ���ܿ��� ( 2013.01.01 ����)                                
            pXLColumn[73] = 35;   // �������� ( 2013.01.01 ����)                             
            pXLColumn[74] = 41;    // �ټӳ�� ( 2013.01.01 ����)    


            //-----------------------------------------

            pXLColumn[75] = 24;   // 34.�����ҵ�(�߰�����)                                            
            pXLColumn[76] = 24;   // 34.�����ҵ�(������)                                             
            pXLColumn[77] = 24;   // 34.�����ҵ�(����(�ջ�) 

            pXLColumn[78] = 24;   // 35.�����ҵ���������                                         
            pXLColumn[79] = 24;   // 36.�ټӿ�������                            
            pXLColumn[80] = 24;   // 37.�����ҵ����ǥ��(27-28-29)    


            //-----------------------------------------
            pXLColumn[81] = 24;   // ����ǥ�ؾȺ�(2012.12.31����)                                       
            pXLColumn[82] = 24;   // ����հ���ǥ��(2012.12.31����)                                   
            pXLColumn[83] = 24;   // ����ջ��⼼��(2012.12.31����) 
            pXLColumn[84] = 24;   // ���⼼�� (2012.12.31����)

            pXLColumn[85] = 30;   // ����ǥ�ؾȺ�(2013.01.01���� )                                               
            pXLColumn[86] = 30;   // ����ǥ��(2013.01.01���� )                                     
            pXLColumn[87] = 30;   // ȯ�����ǥ�� (2013.01.01���� )                                          
            pXLColumn[88] = 30;   // ȯ����⼼�� (2013.01.01���� )  
            pXLColumn[89] = 30;   // ����ջ��⼼�� (2013.01.01���� )            
            pXLColumn[90] = 30;   // ���⼼�� (2013.01.01���� )


            pXLColumn[91] = 36;   // ����ǥ�ؾȺ�(����(�ջ�))                                             
            pXLColumn[92] = 36;   // ����հ���ǥ��(����(�ջ�))
            pXLColumn[93] = 36;   // ȯ�����ǥ��                             
            pXLColumn[94] = 36;   // ȯ����⼼�׾�                                        
            pXLColumn[95] = 36;   // ����ջ��⼼��                            
            pXLColumn[96] = 36;   // ���⼼�� 
            pXLColumn[97] = 24;   // �ⳳ�μ���             
            pXLColumn[98] = 24;   // �Ű��󼼾� 


            //-----------------------------------------
            pXLColumn[99]  = 3;    // �Ű��󼼾� (39) 
            pXLColumn[100] = 7;    // ���ݰ��������
            pXLColumn[101] = 13;   // ����ڵ�Ϲ�ȣ             
            pXLColumn[102] = 18;   // ���¹�ȣ
            pXLColumn[103] = 24;   // �Ա���  
            pXLColumn[104] = 29;   // �����Աݱݾ�                                
            pXLColumn[105] = 34;   // �����޿�           
            pXLColumn[106] = 39;   // �̿������ҵ漼        


            //-----------------------------------------
            pXLColumn[107] = 13;   // �Ű��󼼾�(�ҵ漼)                                
            pXLColumn[108] = 21;   // �Ű��󼼾�(����ҵ漼)
            pXLColumn[109] = 28;   // �Ű��󼼾�(�����Ư����)
            pXLColumn[110] = 36;   // �Ű��󼼾�(�����Ư����)

            pXLColumn[111] = 13;   // �̿������ҵ漼(�ҵ漼)
            pXLColumn[112] = 21;   // �̿������ҵ漼(����ҵ漼)     
            pXLColumn[113] = 28;   // �̿������ҵ漼(�����Ư����) 
            pXLColumn[114] = 36;   // �̿������ҵ漼(�����Ư����) 
            
            pXLColumn[115] = 13;   // ������û¡������(�ҵ漼)                                   
            pXLColumn[116] = 21;   // ������û¡������(����ҵ漼)                   
            pXLColumn[117] = 28;   // ������û¡������(�����Ư����)  
            pXLColumn[118] = 36;   // ������û¡������(�����Ư����) 
            
            pXLColumn[119] = 35;   // �������                       
            pXLColumn[120] = 27;   // ��ǥ�̻�                                    
            //pXLColumn[121] = 12;   // ����������

            pXLColumn[122] = 24;   // 31.�����ҵ�(����(�ջ�))
            pXLColumn[123] = 24;   // 32.�ټӿ�������
            pXLColumn[124] = 24;   // 33.ȯ��޿�
            pXLColumn[125] = 24;   // 34.ȯ��޿�������
            pXLColumn[126] = 24;   // 35.�����ҵ����ǥ��
            pXLColumn[127] = 24;   // 42.ȯ����⼼��
            pXLColumn[128] = 24;   // 43.���⼼��
            pXLColumn[129] = 24;   // 44.�������� ���ϴ� ��������
            pXLColumn[130] = 24;   // 45.Ư��������⼼��
            pXLColumn[131] = 24;   // 46.�ⳳ�� ����
            pXLColumn[132] = 24;   // 47.�Ű��󼼾�
            pXLColumn[133] = 29;   // ���ݰ��� �Ա� �հ� 
            pXLColumn[134] = 2;    // �������� �Ǵ� ����
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

        #region ----- Line Write Method -----

        private int XLLine(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, object pPrintType, string pCourse)
        {
            int vXLine = pXLine; // ������ ������ ǥ�õǴ� �� ��ȣ

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            //System.DateTime vConvertDateTime = new System.DateTime();
            bool IsConvert = false;

            try
            {
                mPrinting.XLActiveSheet("Destination");

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // ���� ����(������1)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���� ����(������2)
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ܱ��� ����(������1) 
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ܱ��� ����(�ܱ���9) 
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��������.
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �������ڵ�
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��� �뵵 ����
                vXLColumnIndex = 15;
                vObject = pPrintType;
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


                // ¡���ǹ�������.
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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


     //---------------------------------------------------------------------------------------------------//
                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // ����ڵ�Ϲ�ȣ 
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���θ�(��ȣ)   
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��ǥ��(����)   
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������(�ּ�) 
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ֹι�ȣ
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ӿ�����
                vGDColumnIndex = pGDColumn[141];
                vXLColumnIndex = pXLColumn[141];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ͼӿ��� ���� ����      
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��������      
                vGDColumnIndex = pGDColumn[142];
                vXLColumnIndex = pXLColumn[142];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����ذ�     
                vGDColumnIndex = pGDColumn[143];
                vXLColumnIndex = pXLColumn[143];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ڹ������� 
                vGDColumnIndex = pGDColumn[144];
                vXLColumnIndex = pXLColumn[144];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ͼӿ��� ������ ����(��������)       
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ӿ�����     
                vGDColumnIndex = pGDColumn[145];
                vXLColumnIndex = pXLColumn[145];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �߰�����     
                vGDColumnIndex = pGDColumn[146];
                vXLColumnIndex = pXLColumn[146];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��Ÿ     
                vGDColumnIndex = pGDColumn[147];
                vXLColumnIndex = pXLColumn[147];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ٹ�ó��-��(��)
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ٹ�ó��- ��(��)
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ٹ�ó��- ��(��)
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ1
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ2
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ3
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����޿�1-����
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿�2-����
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿�3-����
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿� �հ�
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿�1(������� ��)-������ 
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿�2(������� ��)-������ 
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿�3(������� ��)-������ 
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ���������Ͻñ�1
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������Ͻñ�2
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������Ͻñ�3
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������Ͻñ� �հ�
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������� ���������Ͻñ� 1
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������� ���������Ͻñ�2
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������� ���������Ͻñ�3
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������� ���������Ͻñ� �հ�  
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ������1
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������2
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������3
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������ �հ�
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ܰ�1
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ܰ�2
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ܰ�3
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ������ҵ�1
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������ҵ�2
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������ҵ�3
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������ҵ� �հ�
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                //[ ���� �����޿�(��(��)�ٹ���) ]   

                // �Ի���1 - ����      
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����1 - ����
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ���1 - ����
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ���1 - ����
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[86];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������1 - ����
                vGDColumnIndex = pGDColumn[148];
                vXLColumnIndex = pXLColumn[148];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ټӿ���1 - ����
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[87];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // [ ���� �� �����޿�(��(��)�ٹ���) ]
                // �Ի���1 - ���� ��   
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[88];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����1 - ���� ��
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[89];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ���1 - ���� ��
                vGDColumnIndex = pGDColumn[90];
                vXLColumnIndex = pXLColumn[90];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ���1 - ���� ��
                vGDColumnIndex = pGDColumn[91];
                vXLColumnIndex = pXLColumn[91];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������1 - ����
                vGDColumnIndex = pGDColumn[149];
                vXLColumnIndex = pXLColumn[149];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ټӿ���1 - ����
                vGDColumnIndex = pGDColumn[92];
                vXLColumnIndex = pXLColumn[92];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
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

                // [ ���� �����޿�(��(��)�ٹ���) ]

                // �Ի���1 - ����      
                vGDColumnIndex = pGDColumn[93];
                vXLColumnIndex = pXLColumn[93];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����1 - ����
                vGDColumnIndex = pGDColumn[94];
                vXLColumnIndex = pXLColumn[94];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ���1 - ����
                vGDColumnIndex = pGDColumn[95];
                vXLColumnIndex = pXLColumn[95];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ���1 - ����
                vGDColumnIndex = pGDColumn[96];
                vXLColumnIndex = pXLColumn[96];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // [ ���� �� �����޿�(��(��)�ٹ���) ]
                // �Ի���1 - ���� ��   
                vGDColumnIndex = pGDColumn[97];
                vXLColumnIndex = pXLColumn[97];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����1 - ���� ��
                vGDColumnIndex = pGDColumn[98];
                vXLColumnIndex = pXLColumn[98];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ���1 - ���� ��
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ���1 - ���� ��
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

                // �������ݰ��¹�ȣ - ����
                vGDColumnIndex = pGDColumn[150];
                vXLColumnIndex = pXLColumn[150];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �������� �Ͻñ� �Ѽ��ɾ� - ����
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� ������ �հ�� - ����
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� �ҵ��� ���Ծ� - ����
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� �ҵ������ - ����
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� �Ͻñ� - ����
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������� �Ͻñ� �Ѽ��ɾ� - ����
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� ������ �հ�� - ����
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� �ҵ��� ���Ծ� - ����
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� �ҵ������ - ����
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� �Ͻñ� - ����
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������� �Ͻñ� ���� ����� - ���� - ����
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����̿��ݾ�(���������Ͻñ�����) - ���� - ����
                vGDColumnIndex = pGDColumn[151];
                vXLColumnIndex = pXLColumn[151];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������� �����޿��� �޿���- ���� - ����
                vGDColumnIndex = pGDColumn[152];
                vXLColumnIndex = pXLColumn[152];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������� �Ͻñ�- ���� - ����
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ɰ��������޿��� ���� - ����
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ȯ�������ҵ���� ���� - ����
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ȯ�������ҵ����ǥ�� ���� - ����
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ȯ�꿬��հ���ǥ�� ���� - ����
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ȯ�꿬��ջ��⼼�� ���� - ����
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������� �Ͻñ� ���� ����� - ���� - ����
                vGDColumnIndex = pGDColumn[74];
                vXLColumnIndex = pXLColumn[74];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������� �Ͻñ� ���� ����� - ������ - ����
                vGDColumnIndex = pGDColumn[75];
                vXLColumnIndex = pXLColumn[75];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���������� �Ͻñ�- ������ - ����
                vGDColumnIndex = pGDColumn[76];
                vXLColumnIndex = pXLColumn[76];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ɰ��������޿��� ������ - ����
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ȯ�������ҵ���� ������ - ����
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ȯ�������ҵ����ǥ�� ������ - ����
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ȯ�꿬��հ���ǥ�� ������ - ����
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ȯ�꿬��ջ��⼼�� ������ - ����
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������� �Ͻñ� ���� ����� - ������ - ����
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������ݻ���ڸ�
                vGDColumnIndex = pGDColumn[153];
                vXLColumnIndex = pXLColumn[153];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �������ݻ���ڵ�Ϲ�ȣ
                vGDColumnIndex = pGDColumn[154];
                vXLColumnIndex = pXLColumn[154];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �������ݰ��¹�ȣ
                vGDColumnIndex = pGDColumn[155];
                vXLColumnIndex = pXLColumn[155];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���������޿��ݾ�
                vGDColumnIndex = pGDColumn[156];
                vXLColumnIndex = pXLColumn[156];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����������޿��ݾ�
                vGDColumnIndex = pGDColumn[157];
                vXLColumnIndex = pXLColumn[157];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �Ա���
                vGDColumnIndex = pGDColumn[158];
                vXLColumnIndex = pXLColumn[158];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������
                vGDColumnIndex = pGDColumn[159];
                vXLColumnIndex = pXLColumn[159];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����̿�����
                vGDColumnIndex = pGDColumn[160];
                vXLColumnIndex = pXLColumn[160];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �����ݿ��� - ����
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ݿ��� - ���� ��
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ݿ��� - ��
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �����ҵ���� - ����
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ҵ���� - ���� ��
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ҵ���� - ��
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �����ҵ����ǥ�� - ����
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ҵ����ǥ�� - ���� ��
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ҵ����ǥ�� - ��
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ����հ���ǥ�� - ����
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ����հ���ǥ�� - ���� ��
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ����հ���ǥ�� - ��
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ����ջ��⼼�� - ����
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ����ջ��⼼�� - ���� ��
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ����ջ��⼼�� - ��
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ���⼼�� - ����
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���⼼�� - ���� ��
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���⼼�� - ��
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ���⼼�� - ����
                vGDColumnIndex = pGDColumn[121];
                vXLColumnIndex = pXLColumn[121];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���⼼�� - ���� ��
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���⼼�� - ��
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������� - �ҵ漼
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� - ����ҵ漼
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� - �����Ư����
                vGDColumnIndex = pGDColumn[129];
                vXLColumnIndex = pXLColumn[129];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� - ��
                vGDColumnIndex = pGDColumn[130];
                vXLColumnIndex = pXLColumn[130];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �����ٹ����ⳳ�μ��� - �ҵ漼
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ٹ����ⳳ�μ��� - ����ҵ漼
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ٹ����ⳳ�μ��� - �����Ư����
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����ٹ����ⳳ�μ��� - ��
                vGDColumnIndex = pGDColumn[134];
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ������õ¡������ - �ҵ漼
                vGDColumnIndex = pGDColumn[135];
                vXLColumnIndex = pXLColumn[135];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������õ¡������ - ����ҵ漼
                vGDColumnIndex = pGDColumn[136];
                vXLColumnIndex = pXLColumn[136];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������õ¡������ - �����Ư����
                vGDColumnIndex = pGDColumn[137];
                vXLColumnIndex = pXLColumn[137];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������õ¡������ - ��
                vGDColumnIndex = pGDColumn[138];
                vXLColumnIndex = pXLColumn[138];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                
                // ��³�¥
                vGDColumnIndex = pGDColumn[139];
                vXLColumnIndex = pXLColumn[139];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��³�¥
                vGDColumnIndex = pGDColumn[140];
                vXLColumnIndex = pXLColumn[140];
                vObject = pGrid_WITHHOLDING_TAX.GetCellValue(pGridRow, vGDColumnIndex);

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

        #region ----- Line Write Method -----
        private int XLLine2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_2013, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pPrintType, string pPrintType_Desc, string pCourse)
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
                mPrinting.XLActiveSheet("Destination");

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // ���� ����(������1)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���� ����(������2)
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ܱ��� ����(������1) 
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ܱ��� ����(�ܱ���9) 
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��������.
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �������ڵ�
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��� �뵵 ����
                vXLColumnIndex = 15;
                vObject = pPrintType_Desc;
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

                // ¡���ǹ�������.
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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


                //---------------------------------------------------------------------------------------------------//
                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // ����ڵ�Ϲ�ȣ 
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���θ�(��ȣ)   
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��ǥ��(����)   
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ι�ȣ
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������ �ּ�
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ֹι�ȣ 
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ӿ����� 
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // Ȯ���޿��� �������� ���� ������  
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������ ��¥ 
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // �ͼӿ��� ���� ���� 
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��������.
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����ذ�
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ڹ�������
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ͼӿ��� ������ ����(��������). 
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ӿ�����
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                //�߰�����
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                //��Ÿ
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                //----�����޿���Ȳ----

                // �ٹ�ó�� ( �߰�����)
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ٹ�ó�� (������)  
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ ( �߰�����)  
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ (������)          
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����޿� (�߰�����)
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿� (������)  
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿� (����)
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ����� �����޿� (�߰�����)       
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ����� �����޿� (������)    
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ����� �����޿� (����)   
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ������� �����޿� (�߰�����)
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������� ������ (������)   
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������� ������ (����) 
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                //-------�ټӿ���------------------
                // �Ի��� ( �߰����� �ټӿ���) 
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����/��������� ( �߰����� �ټӿ���)  
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����� ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������ ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ܿ��� ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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


                // �������� ( �߰����� �ټӿ���)  
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �Ի���(������)       
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����/��������� ( ������)   
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����� ( ������ )  
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������ ( ������ )
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� ( ������)
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ��� ( ������) 
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������� ( ������)    
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ټӿ��� ( ������)  
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
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

                // �Ի���(����(�ջ�))  
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����/��������� (����(�ջ�))   
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����� (����(�ջ�))
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� (����(�ջ�))
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ��� (����(�ջ�))  
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� (����(�ջ�))
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ߺ����� (����(�ջ�)) 
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ټӿ��� (����(�ջ�)) 
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
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

                // ����� (2012.12.31 ����)  
                vGDColumnIndex = pGDColumn[63];
                vXLColumnIndex = pXLColumn[63];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����  (2012.12.31 ����) 
                vGDColumnIndex = pGDColumn[64];
                vXLColumnIndex = pXLColumn[64];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� (2012.12.31 ����)   
                vGDColumnIndex = pGDColumn[65];
                vXLColumnIndex = pXLColumn[65];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                // ���ܿ��� (2012.12.31 ����)
                vGDColumnIndex = pGDColumn[66];
                vXLColumnIndex = pXLColumn[66];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� (2012.12.31 ����) 
                vGDColumnIndex = pGDColumn[67];
                vXLColumnIndex = pXLColumn[67];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ټӳ�� (2012.12.31 ����)
                vGDColumnIndex = pGDColumn[68];
                vXLColumnIndex = pXLColumn[68];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
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

                // ����� ( 2013.01.01 ����)    
                vGDColumnIndex = pGDColumn[69];
                vXLColumnIndex = pXLColumn[69];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����  ( 2013.01.01 ����)
                vGDColumnIndex = pGDColumn[70];
                vXLColumnIndex = pXLColumn[70];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� ( 2013.01.01 ����) 
                vGDColumnIndex = pGDColumn[71];
                vXLColumnIndex = pXLColumn[71];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ��� ( 2013.01.01 ����)   
                vGDColumnIndex = pGDColumn[72];
                vXLColumnIndex = pXLColumn[72];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� ( 2013.01.01 ����) 
                vGDColumnIndex = pGDColumn[73];
                vXLColumnIndex = pXLColumn[73];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //�ټӳ�� ( 2013.01.01 ����)
                vGDColumnIndex = pGDColumn[74];
                vXLColumnIndex = pXLColumn[74];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
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

                //----------------------- ���������� ���� ����� ----------------------------

                // 27.�����ҵ�(17)
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 28.�ټӿ�������
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 29.ȯ��޿�(27-28) * 12�� / ���� �ټӿ���   
                vGDColumnIndex = pGDColumn[124];
                vXLColumnIndex = pXLColumn[124];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 30.ȯ��޿�������      
                vGDColumnIndex = pGDColumn[125];
                vXLColumnIndex = pXLColumn[125];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 31.�����ҵ����ǥ��
                vGDColumnIndex = pGDColumn[126];
                vXLColumnIndex = pXLColumn[126];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 32.ȯ����⼼��(31 * ����)
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 33.���⼼��(31*����ټӿ��� / 12��)
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 34.�����ҵ�(17)
                vGDColumnIndex = pGDColumn[77];
                vXLColumnIndex = pXLColumn[77];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 35.�����ҵ���������
                vGDColumnIndex = pGDColumn[78];
                vXLColumnIndex = pXLColumn[78];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 36.�ټӿ�������
                vGDColumnIndex = pGDColumn[79];
                vXLColumnIndex = pXLColumn[79];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 37.�����ҵ����ǥ��(34-35-36)
                vGDColumnIndex = pGDColumn[80];
                vXLColumnIndex = pXLColumn[80];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 38.����ǥ�ؾȺ�(37*���ټӿ���/����ټӿ���) - 2012.12.31 ����   
                vGDColumnIndex = pGDColumn[81];
                vXLColumnIndex = pXLColumn[81];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 38.����ǥ�ؾȺ�(37*���ټӿ���/����ټӿ���) - 2013.01.01 ����   
                vGDColumnIndex = pGDColumn[85];
                vXLColumnIndex = pXLColumn[85];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 38.����ǥ�ؾȺ�(37*���ټӿ���/����ټӿ���) - �հ�  
                vGDColumnIndex = pGDColumn[91];
                vXLColumnIndex = pXLColumn[91];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 39.����հ���ǥ��(38/���ټӿ���) - 2012.12.31���� 
                vGDColumnIndex = pGDColumn[82];
                vXLColumnIndex = pXLColumn[82];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 39.����հ���ǥ��(38/���ټӿ���) - 2013.01.01���� 
                vGDColumnIndex = pGDColumn[86];
                vXLColumnIndex = pXLColumn[86];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 39.����հ���ǥ��(38/���ټӿ���) - �հ�
                vGDColumnIndex = pGDColumn[92];
                vXLColumnIndex = pXLColumn[92];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);
                vConvertDecimal = iString.ISDecimaltoZero(vObject, 0);
                if (vConvertDecimal != 0)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //40.ȯ�����ǥ��(39*5��)-2013.01.01���� 
                vGDColumnIndex = pGDColumn[87];
                vXLColumnIndex = pXLColumn[87];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                //40.ȯ�����ǥ��(39*5��)-�հ� 
                vGDColumnIndex = pGDColumn[93];
                vXLColumnIndex = pXLColumn[93];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 41.ȯ����⼼��(40*����)-2013.01.01����
                vGDColumnIndex = pGDColumn[88];
                vXLColumnIndex = pXLColumn[88];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 41.ȯ����⼼��(40*����)-�հ�
                vGDColumnIndex = pGDColumn[94];
                vXLColumnIndex = pXLColumn[94];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 42.����ջ��⼼��-2012.12.31���� 
                vGDColumnIndex = pGDColumn[83];
                vXLColumnIndex = pXLColumn[83];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 42.����ջ��⼼��-2013.01.01 ����  
                vGDColumnIndex = pGDColumn[89];
                vXLColumnIndex = pXLColumn[89];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 42.����ջ��⼼��-�հ�
                vGDColumnIndex = pGDColumn[95];
                vXLColumnIndex = pXLColumn[95];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 43.���⼼��(42*���ټӿ���)-2012.12.31 ���� 
                vGDColumnIndex = pGDColumn[84];
                vXLColumnIndex = pXLColumn[84];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 43.���⼼��(42*���ټӿ���)-2013.01.01 ����  
                vGDColumnIndex = pGDColumn[90];
                vXLColumnIndex = pXLColumn[90];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // 43.���⼼��(42*���ټӿ���)-�հ� 
                vGDColumnIndex = pGDColumn[96];
                vXLColumnIndex = pXLColumn[96];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 44.�������� ���ϴ� �������� 
                vGDColumnIndex = pGDColumn[129];
                vXLColumnIndex = pXLColumn[129];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
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
                //45.�����ҵ漼 ���⼼�� 
                vGDColumnIndex = pGDColumn[130];
                vXLColumnIndex = pXLColumn[130];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //46.�ⳳ��(�Ǵ� ������̿�) ���� 
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //47.�Ű��󼼾�(45-46)
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //48.�Ű��� ���� 
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ݰ�������
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ 
                vGDColumnIndex = pGDColumn[101];
                vXLColumnIndex = pXLColumn[101];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                vGDColumnIndex = pGDColumn[102];
                vXLColumnIndex = pXLColumn[102];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �Ա���  
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����Աݱݾ� 
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿�   
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �̿������ҵ漼   
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 52.���ݰ��� �Աݱݾ� �հ�  
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 53.�Ű��󼼾�(�ҵ漼)  
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �Ű��󼼾�(����ҵ漼)
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �Ű��󼼾�(�����Ư����)
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �Ű��󼼾�(��)
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �̿������ҵ漼(�ҵ漼)  
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �̿������ҵ漼(����ҵ漼)  
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �̿������ҵ漼(�����Ư����)    
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �̿������ҵ漼(��)
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // ������û¡������(�ҵ漼)
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������û¡������(����ҵ漼)  
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������û¡������(�����Ư����)  
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������û¡������(��)  
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��ǥ�̻�    
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                // ������ ��,�ҵ��� �������̸� �ҵ��ڸ� �μ�
                if (pPrintType == "1")
                {
                    vGDColumnIndex = pGDColumn[12];
                }
                else
                {
                    vGDColumnIndex = pGDColumn[134];
                }
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

        private int XLLine3(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_2013, int pGridRow, int pXLine, int[] pGDColumn, int[] pXLColumn, string pPrintType, string pPrintType_Desc, string pCourse)
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
                mPrinting.XLActiveSheet("Destination");

                //-------------------------------------------------------------------
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // ���� ����(������1)
                vGDColumnIndex = pGDColumn[0];
                vXLColumnIndex = pXLColumn[0];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���� ����(������2)
                vGDColumnIndex = pGDColumn[1];
                vXLColumnIndex = pXLColumn[1];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ܱ��� ����(������1) 
                vGDColumnIndex = pGDColumn[2];
                vXLColumnIndex = pXLColumn[2];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ܱ��� ����(�ܱ���9) 
                vGDColumnIndex = pGDColumn[3];
                vXLColumnIndex = pXLColumn[3];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                ///-------�������������ڿ���
               
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                // ��������.
                vGDColumnIndex = pGDColumn[4];
                vXLColumnIndex = pXLColumn[4];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �������ڵ�
                vGDColumnIndex = pGDColumn[5];
                vXLColumnIndex = pXLColumn[5];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��� �뵵 ����
                vXLColumnIndex = 15;
                vObject = pPrintType_Desc;
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

                // ¡���ǹ�������.
                vGDColumnIndex = pGDColumn[6];
                vXLColumnIndex = pXLColumn[6];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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


                //---------------------------------------------------------------------------------------------------//
                //-------------------------------------------------------------------
                vXLine = vXLine + 3;
                //-------------------------------------------------------------------

                // ����ڵ�Ϲ�ȣ 
                vGDColumnIndex = pGDColumn[7];
                vXLColumnIndex = pXLColumn[7];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���θ�(��ȣ)   
                vGDColumnIndex = pGDColumn[8];
                vXLColumnIndex = pXLColumn[8];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��ǥ��(����)   
                vGDColumnIndex = pGDColumn[9];
                vXLColumnIndex = pXLColumn[9];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ι�ȣ
                vGDColumnIndex = pGDColumn[10];
                vXLColumnIndex = pXLColumn[10];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������ �ּ�
                vGDColumnIndex = pGDColumn[11];
                vXLColumnIndex = pXLColumn[11];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����
                vGDColumnIndex = pGDColumn[12];
                vXLColumnIndex = pXLColumn[12];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ֹι�ȣ 
                vGDColumnIndex = pGDColumn[13];
                vXLColumnIndex = pXLColumn[13];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                vGDColumnIndex = pGDColumn[14];
                vXLColumnIndex = pXLColumn[14];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ӿ����� 
                vGDColumnIndex = pGDColumn[15];
                vXLColumnIndex = pXLColumn[15];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // Ȯ���޿��� �������� ���� ������  
                vGDColumnIndex = pGDColumn[16];
                vXLColumnIndex = pXLColumn[16];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������ ��¥ 
                vGDColumnIndex = pGDColumn[17];
                vXLColumnIndex = pXLColumn[17];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                vXLine = vXLine + 2;
                //-------------------------------------------------------------------

                // �ͼӿ��� ���� ���� 
                vGDColumnIndex = pGDColumn[18];
                vXLColumnIndex = pXLColumn[18];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��������.
                vGDColumnIndex = pGDColumn[20];
                vXLColumnIndex = pXLColumn[20];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����ذ�
                vGDColumnIndex = pGDColumn[21];
                vXLColumnIndex = pXLColumn[21];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ڹ�������
                vGDColumnIndex = pGDColumn[22];
                vXLColumnIndex = pXLColumn[22];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ͼӿ��� ������ ����(��������). 
                vGDColumnIndex = pGDColumn[19];
                vXLColumnIndex = pXLColumn[19];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ӿ�����
                vGDColumnIndex = pGDColumn[23];
                vXLColumnIndex = pXLColumn[23];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                //�߰�����
                vGDColumnIndex = pGDColumn[24];
                vXLColumnIndex = pXLColumn[24];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                //��Ÿ
                vGDColumnIndex = pGDColumn[25];
                vXLColumnIndex = pXLColumn[25];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                //----�����޿���Ȳ----

                // �ٹ�ó�� ( �߰�����)
                vGDColumnIndex = pGDColumn[26];
                vXLColumnIndex = pXLColumn[26];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ٹ�ó�� (������)  
                vGDColumnIndex = pGDColumn[27];
                vXLColumnIndex = pXLColumn[27];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ ( �߰�����)  
                vGDColumnIndex = pGDColumn[28];
                vXLColumnIndex = pXLColumn[28];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ (������)          
                vGDColumnIndex = pGDColumn[29];
                vXLColumnIndex = pXLColumn[29];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����޿� (�߰�����)
                vGDColumnIndex = pGDColumn[30];
                vXLColumnIndex = pXLColumn[30];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿� (������)  
                vGDColumnIndex = pGDColumn[31];
                vXLColumnIndex = pXLColumn[31];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿� (����)
                vGDColumnIndex = pGDColumn[32];
                vXLColumnIndex = pXLColumn[32];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ����� �����޿� (�߰�����)       
                vGDColumnIndex = pGDColumn[33];
                vXLColumnIndex = pXLColumn[33];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ����� �����޿� (������)    
                vGDColumnIndex = pGDColumn[34];
                vXLColumnIndex = pXLColumn[34];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ����� �����޿� (����)   
                vGDColumnIndex = pGDColumn[35];
                vXLColumnIndex = pXLColumn[35];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // ������� �����޿� (�߰�����)
                vGDColumnIndex = pGDColumn[36];
                vXLColumnIndex = pXLColumn[36];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������� ������ (������)   
                vGDColumnIndex = pGDColumn[37];
                vXLColumnIndex = pXLColumn[37];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������� ������ (����) 
                vGDColumnIndex = pGDColumn[38];
                vXLColumnIndex = pXLColumn[38];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                //-------�ټӿ���------------------
                // �Ի��� ( �߰����� �ټӿ���) 
                vGDColumnIndex = pGDColumn[39];
                vXLColumnIndex = pXLColumn[39];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����/��������� ( �߰����� �ټӿ���)  
                vGDColumnIndex = pGDColumn[40];
                vXLColumnIndex = pXLColumn[40];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����� ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[41];
                vXLColumnIndex = pXLColumn[41];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������ ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[42];
                vXLColumnIndex = pXLColumn[42];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[43];
                vXLColumnIndex = pXLColumn[43];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ���ܿ��� ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[44];
                vXLColumnIndex = pXLColumn[44];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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


                // �������� ( �߰����� �ټӿ���)  
                vGDColumnIndex = pGDColumn[45];
                vXLColumnIndex = pXLColumn[45];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� ( �߰����� �ټӿ���)
                vGDColumnIndex = pGDColumn[46];
                vXLColumnIndex = pXLColumn[46];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �Ի���(������)       
                vGDColumnIndex = pGDColumn[47];
                vXLColumnIndex = pXLColumn[47];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����/��������� ( ������)   
                vGDColumnIndex = pGDColumn[48];
                vXLColumnIndex = pXLColumn[48];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����� ( ������ )  
                vGDColumnIndex = pGDColumn[49];
                vXLColumnIndex = pXLColumn[49];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ������ ( ������ )
                vGDColumnIndex = pGDColumn[50];
                vXLColumnIndex = pXLColumn[50];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� ( ������)
                vGDColumnIndex = pGDColumn[51];
                vXLColumnIndex = pXLColumn[51];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ��� ( ������) 
                vGDColumnIndex = pGDColumn[52];
                vXLColumnIndex = pXLColumn[52];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������� ( ������)    
                vGDColumnIndex = pGDColumn[53];
                vXLColumnIndex = pXLColumn[53];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ټӿ��� ( ������)  
                vGDColumnIndex = pGDColumn[54];
                vXLColumnIndex = pXLColumn[54];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
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

                // �Ի���(����(�ջ�))  
                vGDColumnIndex = pGDColumn[55];
                vXLColumnIndex = pXLColumn[55];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����/��������� (����(�ջ�))   
                vGDColumnIndex = pGDColumn[56];
                vXLColumnIndex = pXLColumn[56];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����� (����(�ջ�))
                vGDColumnIndex = pGDColumn[57];
                vXLColumnIndex = pXLColumn[57];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �ټӿ��� (����(�ջ�))
                vGDColumnIndex = pGDColumn[58];
                vXLColumnIndex = pXLColumn[58];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ܿ��� (����(�ջ�))  
                vGDColumnIndex = pGDColumn[59];
                vXLColumnIndex = pXLColumn[59];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �������� (����(�ջ�))
                vGDColumnIndex = pGDColumn[60];
                vXLColumnIndex = pXLColumn[60];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ߺ����� (����(�ջ�)) 
                vGDColumnIndex = pGDColumn[61];
                vXLColumnIndex = pXLColumn[61];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �ټӿ��� (����(�ջ�)) 
                vGDColumnIndex = pGDColumn[62];
                vXLColumnIndex = pXLColumn[62];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertDecimal);
                    if (vConvertString == "0")
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
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

               

                //----------------------- ���������� ���� ����� ----------------------------
                
                // 27.�����ҵ�(17)
                vGDColumnIndex = pGDColumn[122];
                vXLColumnIndex = pXLColumn[122];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 28.�ټӿ�������
                vGDColumnIndex = pGDColumn[123];
                vXLColumnIndex = pXLColumn[123];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 29.ȯ��޿�(27-28) * 12�� / ���� �ټӿ���   
                vGDColumnIndex = pGDColumn[124];
                vXLColumnIndex = pXLColumn[124];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 30.ȯ��޿�������      
                vGDColumnIndex = pGDColumn[125];
                vXLColumnIndex = pXLColumn[125];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 31.�����ҵ����ǥ��
                vGDColumnIndex = pGDColumn[126];
                vXLColumnIndex = pXLColumn[126];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 32.ȯ����⼼��(31 * ����)
                vGDColumnIndex = pGDColumn[127];
                vXLColumnIndex = pXLColumn[127];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 33.���⼼��(31*����ټӿ��� / 12��)
                vGDColumnIndex = pGDColumn[128];
                vXLColumnIndex = pXLColumn[128];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
                //// 34.�����ҵ�(17)
                //vGDColumnIndex = pGDColumn[77];
                //vXLColumnIndex = pXLColumn[77];
                //vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);
                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}

                ////-------------------------------------------------------------------
                //vXLine = vXLine + 2;
                ////-------------------------------------------------------------------
                //// 35.�����ҵ���������
                //vGDColumnIndex = pGDColumn[78];
                //vXLColumnIndex = pXLColumn[78];
                //vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////-------------------------------------------------------------------
                //vXLine = vXLine + 1;
                ////-------------------------------------------------------------------

                ////45.�����ҵ漼 ���⼼�� 
                //vGDColumnIndex = pGDColumn[130];
                //vXLColumnIndex = pXLColumn[130];
                //vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////-------------------------------------------------------------------
                //vXLine = vXLine + 1;
                ////-------------------------------------------------------------------

                ////45.�����ҵ漼 ���⼼�� 
                //vGDColumnIndex = pGDColumn[130];
                //vXLColumnIndex = pXLColumn[130];
                //vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                //IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                //}
                ////-------------------------------------------------------------------
                //vXLine = vXLine + 1;
                ////-------------------------------------------------------------------
 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //46.�ⳳ��(�Ǵ� ������̿�) ���� 
                vGDColumnIndex = pGDColumn[131];
                vXLColumnIndex = pXLColumn[131];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                //47.�Ű��󼼾�(45-46)
                vGDColumnIndex = pGDColumn[132];
                vXLColumnIndex = pXLColumn[132];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }


                //-------------------------------------------------------------------
                vXLine = vXLine + 5;
                //-------------------------------------------------------------------
                //48.�Ű��� ���� 
                vGDColumnIndex = pGDColumn[99];
                vXLColumnIndex = pXLColumn[99];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ���ݰ�������
                vGDColumnIndex = pGDColumn[100];
                vXLColumnIndex = pXLColumn[100];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ����ڵ�Ϲ�ȣ 
                vGDColumnIndex = pGDColumn[101];
                vXLColumnIndex = pXLColumn[101];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                vGDColumnIndex = pGDColumn[102];
                vXLColumnIndex = pXLColumn[102];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �Ա���  
                vGDColumnIndex = pGDColumn[103];
                vXLColumnIndex = pXLColumn[103];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // �����Աݱݾ� 
                vGDColumnIndex = pGDColumn[104];
                vXLColumnIndex = pXLColumn[104];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �����޿�   
                vGDColumnIndex = pGDColumn[105];
                vXLColumnIndex = pXLColumn[105];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �̿������ҵ漼   
                vGDColumnIndex = pGDColumn[106];
                vXLColumnIndex = pXLColumn[106];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 52.���ݰ��� �Աݱݾ� �հ�  
                vGDColumnIndex = pGDColumn[133];
                vXLColumnIndex = pXLColumn[133];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // 53.�Ű��󼼾�(�ҵ漼)  
                vGDColumnIndex = pGDColumn[107];
                vXLColumnIndex = pXLColumn[107];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �Ű��󼼾�(����ҵ漼)
                vGDColumnIndex = pGDColumn[108];
                vXLColumnIndex = pXLColumn[108];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �Ű��󼼾�(�����Ư����)
                vGDColumnIndex = pGDColumn[109];
                vXLColumnIndex = pXLColumn[109];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �Ű��󼼾�(��)
                vGDColumnIndex = pGDColumn[110];
                vXLColumnIndex = pXLColumn[110];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �̿������ҵ漼(�ҵ漼)  
                vGDColumnIndex = pGDColumn[111];
                vXLColumnIndex = pXLColumn[111];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �̿������ҵ漼(����ҵ漼)  
                vGDColumnIndex = pGDColumn[112];
                vXLColumnIndex = pXLColumn[112];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �̿������ҵ漼(�����Ư����)    
                vGDColumnIndex = pGDColumn[113];
                vXLColumnIndex = pXLColumn[113];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // �̿������ҵ漼(��)
                vGDColumnIndex = pGDColumn[114];
                vXLColumnIndex = pXLColumn[114];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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
                // ������û¡������(�ҵ漼)
                vGDColumnIndex = pGDColumn[115];
                vXLColumnIndex = pXLColumn[115];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������û¡������(����ҵ漼)  
                vGDColumnIndex = pGDColumn[116];
                vXLColumnIndex = pXLColumn[116];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������û¡������(�����Ư����)  
                vGDColumnIndex = pGDColumn[117];
                vXLColumnIndex = pXLColumn[117];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                    mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                }

                // ������û¡������(��)  
                vGDColumnIndex = pGDColumn[118];
                vXLColumnIndex = pXLColumn[118];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:#,###,###,###,###,###}", vConvertDecimal);
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

                // �������
                vGDColumnIndex = pGDColumn[119];
                vXLColumnIndex = pXLColumn[119];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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

                // ��ǥ�̻�    
                vGDColumnIndex = pGDColumn[120];
                vXLColumnIndex = pXLColumn[120];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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
                // ������ ��,�ҵ��� �������̸� �ҵ��ڸ� �μ�
                if (pPrintType == "1")
                {
                    vGDColumnIndex = pGDColumn[12];
                }
                else
                {
                    vGDColumnIndex = pGDColumn[134];
                }
                vXLColumnIndex = pXLColumn[134];
                vObject = pGrid_PRINT_2013.GetCellValue(pGridRow, vGDColumnIndex);

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


        #region ----- Excel Write WithholdingTax  Method ----

        public int WriteWithholdingTax(string pPrint_Type, string pSaveFileName, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_WITHHOLDING_TAX, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid_PRINT_2013, string pPrintType, string pPrintType_Desc)
        {
            string vMessageText = string.Empty;
            bool isOpen = XLFileOpen();
            mCopyLineSUM = 1;
            mPageNumber = 0;

            int[] vGDColumn;
            int[] vXLColumn;

            int vTotalRow = pGrid_WITHHOLDING_TAX.RowCount;
            int vTotalRow2  = pGrid_PRINT_2013.RowCount;
            string vRetire_Year = pGrid_PRINT_2013.GetCellValue("FINAL_RETIRE_DATE").ToString();


            int vRowCount = 0;

            int vPrintingLine = 0;

            int vSecondPrinting = 9;
            int vCountPrinting = 0;

            if (iString.ISNumtoZero(vRetire_Year) < 2013)
            {
                SetArray1(pGrid_WITHHOLDING_TAX, out vGDColumn, out vXLColumn);

                for (int vRow = 0; vRow < vTotalRow; vRow++)
                {
                    vRowCount++;
                    pGrid_WITHHOLDING_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                    mAppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();

                    if (isOpen == true)
                    {
                        vCountPrinting++;

                        mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                        pGrid_WITHHOLDING_TAX.CurrentCellMoveTo(vRow, 0);
                        pGrid_WITHHOLDING_TAX.Focus();
                        pGrid_WITHHOLDING_TAX.CurrentCellActivate(vRow, 0);

                        // �����ҵ��õ¡��������/��������
                        vPrintingLine = XLLine(pGrid_WITHHOLDING_TAX, vRow, vPrintingLine, vGDColumn, vXLColumn, pPrintType, "SRC_TAB1");

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
            }
            else if(iString.ISNumtoZero(vRetire_Year) < 2020)
            {
                SetArray2(pGrid_PRINT_2013, out vGDColumn, out vXLColumn);

                for (int vRow = 0; vRow < vTotalRow; vRow++)
                {
                    vRowCount++;
                    pGrid_WITHHOLDING_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                    mAppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();

                    if (isOpen == true)
                    {
                        vCountPrinting++;

                        mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                        pGrid_PRINT_2013.CurrentCellMoveTo(vRow, 0);
                        pGrid_PRINT_2013.Focus();
                        pGrid_PRINT_2013.CurrentCellActivate(vRow, 0);

                        // �����ҵ��õ¡��������/��������
                        vPrintingLine = XLLine2(pGrid_PRINT_2013, vRow, vPrintingLine, vGDColumn, vXLColumn, pPrintType, pPrintType_Desc, "SRC_TAB1");

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
            }
            else
            {
                SetArray2(pGrid_PRINT_2013, out vGDColumn, out vXLColumn);

                for (int vRow = 0; vRow < vTotalRow; vRow++)
                {
                    vRowCount++;
                    pGrid_WITHHOLDING_TAX.Cursor = System.Windows.Forms.Cursors.WaitCursor;

                    vMessageText = string.Format("Printing : {0}/{1}", vRowCount, vTotalRow);
                    mAppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();

                    if (isOpen == true)
                    {
                        vCountPrinting++;

                        mCopyLineSUM = CopyAndPaste(mPrinting, mCopyLineSUM, "SRC_TAB1");
                        vPrintingLine = (mCopyLineSUM - mIncrementCopyMAX) + (mPrintingLineSTART - 1);

                        pGrid_PRINT_2013.CurrentCellMoveTo(vRow, 0);
                        pGrid_PRINT_2013.Focus();
                        pGrid_PRINT_2013.CurrentCellActivate(vRow, 0);

                        // �����ҵ��õ¡��������/��������
                        vPrintingLine = XLLine3(pGrid_PRINT_2013, vRow, vPrintingLine, vGDColumn, vXLColumn, pPrintType, pPrintType_Desc, "SRC_TAB1");

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
            }
            //SAVE("RETIRE.XLS");
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

            if (pCourse == "SRC_TAB1")
            {
                pPrinting.XLActiveSheet("SourceTab1");
            }            

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