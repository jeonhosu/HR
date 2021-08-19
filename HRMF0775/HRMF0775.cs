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

namespace HRMF0775
{
    public partial class HRMF0775 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0775()
        {
            InitializeComponent();
        }

        public HRMF0775(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

        //업체
        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            IDA_RESIDENT_ETC.Fill();
            IGR_RESIDENT_BUSINESS.Focus();
        }

        private void SetCommonParameter(object P_GROUP_CODE, object P_ENABLED_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", P_GROUP_CODE);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", P_ENABLED_YN);
        }

        private void Set_Tot_Payment_Amount(decimal pPayment_amount, decimal pPayment_etc_amount)
        {
            IGR_RESIDENT_INCOME.SetCellValue("TOT_PAYMENT_AMOUNT", pPayment_amount + pPayment_etc_amount); 

            //필요경비 동기화//
            Set_Exp_Amount(iConv.ISDecimaltoZero(IGR_RESIDENT_INCOME.GetCellValue("EXP_RATE")));
        }

        private void Set_Exp_Amount(decimal pExp_Rate)
        {
            decimal vExp_amount = 0;
            decimal vTot_payment_amount = iConv.ISDecimaltoZero(IGR_RESIDENT_INCOME.GetCellValue("TOT_PAYMENT_AMOUNT"), 0);
            vExp_amount = Math.Truncate(vTot_payment_amount * (pExp_Rate / 100));

            IGR_RESIDENT_INCOME.SetCellValue("EXP_AMOUNT", vExp_amount);

            //소득금액 동기화 //
            IGR_RESIDENT_INCOME.SetCellValue("INCOME_AMOUNT", vTot_payment_amount - vExp_amount);
            Set_Tot_Income_Amount(vTot_payment_amount - vExp_amount,  
                                    iConv.ISDecimaltoZero(IGR_RESIDENT_INCOME.GetCellValue("INCOME_ETC_AMOUNT")));             
        }

        private void Set_Tot_Income_Amount(decimal pIncom_amount, decimal pIncom_etc_amount)
        {
            IGR_RESIDENT_INCOME.SetCellValue("TOT_INCOME_AMOUNT", pIncom_amount + pIncom_etc_amount);

            //세금 동기화//
            Set_Income_Tax_Amount(IGR_RESIDENT_INCOME.GetCellValue("TAX_RATE"));
        }

        private void Set_Income_Tax_Amount(object pTax_Rate)
        {
            IDC_RESIDENT_ETC_INCOME_TAX_AMT_P.SetCommandParamValue("P_STD_DATE", IGR_RESIDENT_INCOME.GetCellValue("PAY_DATE"));
            IDC_RESIDENT_ETC_INCOME_TAX_AMT_P.SetCommandParamValue("P_TOT_INCOME_AMOUNT", IGR_RESIDENT_INCOME.GetCellValue("TOT_INCOME_AMOUNT"));
            IDC_RESIDENT_ETC_INCOME_TAX_AMT_P.SetCommandParamValue("P_TAX_RATE", pTax_Rate);
            IDC_RESIDENT_ETC_INCOME_TAX_AMT_P.ExecuteNonQuery();
            object vIncom_tax_amount = IDC_RESIDENT_ETC_INCOME_TAX_AMT_P.GetCommandParamValue("O_INCOME_TAX_AMT");                         
            IGR_RESIDENT_INCOME.SetCellValue("INCOME_TAX_AMT", vIncom_tax_amount); 
            
            //지방소득세//
            Set_Local_Tax_Amount(vIncom_tax_amount);
        }

        private void Set_Local_Tax_Amount(object pIncome_Tax_Amount)
        {
            IDC_RESIDENT_ETC_LOCAL_TAX_AMT_P.SetCommandParamValue("P_STD_DATE", IGR_RESIDENT_INCOME.GetCellValue("PAY_DATE"));
            IDC_RESIDENT_ETC_LOCAL_TAX_AMT_P.SetCommandParamValue("P_INCOME_TAX_AMT", pIncome_Tax_Amount);
            IDC_RESIDENT_ETC_LOCAL_TAX_AMT_P.ExecuteNonQuery();
            object vLocal_tax_amount = IDC_RESIDENT_ETC_LOCAL_TAX_AMT_P.GetCommandParamValue("O_LOCAL_TAX_AMT");
             
            IGR_RESIDENT_INCOME.SetCellValue("LOCAL_TAX_AMT", vLocal_tax_amount);

            Set_Tot_Dedction_Amount(pIncome_Tax_Amount, vLocal_tax_amount);
        }

        private void Set_Tot_Dedction_Amount(object pIncome_Tax_Amount, object pLocal_Tax_Amount)
        {
            decimal vTot_dedction_amount = iConv.ISDecimaltoZero(pIncome_Tax_Amount, 0) + 
                                            iConv.ISDecimaltoZero(pLocal_Tax_Amount, 0);
            IGR_RESIDENT_INCOME.SetCellValue("TOTAL_DED_AMT", vTot_dedction_amount);

            Set_Real_Amount(iConv.ISDecimaltoZero(IGR_RESIDENT_INCOME.GetCellValue("TOT_PAYMENT_AMOUNT")),
                            vTot_dedction_amount);
        }

        private void Set_Real_Amount(decimal pTot_payment_amount, decimal pTot_dedction_amount)
        {
            IGR_RESIDENT_INCOME.SetCellValue("REAL_AMT", pTot_payment_amount - pTot_dedction_amount); 
        }

        #endregion;

        #region ----- 주민번호 체크 -----

        private object REPRE_NUM_Check(object pRepre_num)
        {
            object Check_YN = "N";
            if (iConv.ISNull(pRepre_num) == string.Empty)
            {
                return Check_YN;
            }
                        
            idcREPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", pRepre_num);
            idcREPRE_NUM_CHECK.ExecuteNonQuery();
            Check_YN = idcREPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE");
            return Check_YN;
        }

        #endregion
        
        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            try
            {                
                switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
                {
                    case ISUtil.Enum.TerritoryLanguage.Default:
                        mPrompt = pEdit.PromptTextElement[mIDX].Default;
                        break;
                    case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                        mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                        break;
                    case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                        mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                        break;
                    case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                        mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                        break;
                    case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                        mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                        break;
                    case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                        mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                        break;
                }
            }
            catch
            {
            }
            return mPrompt;
        }

        #endregion;


        #region ----- XL Print 1 Methods -----

        private void Print_Doc(string pOutput_Type)
        {
            DialogResult dlgRESULT;
            HRMF0775_PRINT vHRMF0775_PRINT = new HRMF0775_PRINT(isAppInterfaceAdv1.AppInterface);
            dlgRESULT = vHRMF0775_PRINT.ShowDialog();
            if (dlgRESULT == DialogResult.Cancel)
            {
                return;
            }
            string vPrint_1_YN = vHRMF0775_PRINT.Print_1_YN;
            string vPrint_2_YN = vHRMF0775_PRINT.Print_2_YN;

            if (vPrint_1_YN == "Y")
            {
                XLPrinting1(pOutput_Type, "1");
            }
            if (vPrint_2_YN == "Y")
            {
                XLPrinting1(pOutput_Type, "2");
            }
        }

        private void XLPrinting1(string pOutput_Type, object pPrint_Type)
        {          
            string vMessageText = string.Empty;
            string vFilePath = string.Empty;
            string vSaveFileName = string.Empty;
            int vPageNumber = 0;
            int vCountRow = 0;
            

            // 데이터 조회.
            IDA_PRINT_RESIDENT_WH_ETC.SetSelectParamValue("P_PRINT_TYPE", pPrint_Type);
            IDA_PRINT_RESIDENT_WH_ETC.Fill();
            vCountRow = IDA_PRINT_RESIDENT_WH_ETC.OraSelectData.Rows.Count;
                        
            if (vCountRow < 1)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10386"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (pOutput_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = "Income_etc";

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.DefaultExt = "xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vFilePath = saveFileDialog1.FileName;
                    vSaveFileName = vFilePath;

                    System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);
                    try
                    {
                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }
                    }
                    catch (Exception EX)
                    {
                        MessageBoxAdv.Show(EX.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            //원화 인쇄//
            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);
            try
            {
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0775_001.xlsx";
                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                bool isOpen = xlPrinting.XLFileOpen();
                //-------------------------------------------------------------------------------------

                if (isOpen == true)
                {
                    vMessageText = string.Format(" Printing Starting...");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);

                    vPageNumber = xlPrinting.ExcelWrite(IDA_PRINT_RESIDENT_WH_ETC, IDA_PRINT_RESIDENT_INCOME_ETC);

                    if (pOutput_Type == "PRINT")
                    {
                        //[PRINTING]
                        xlPrinting.Printing(1, vPageNumber); //시작 페이지 번호, 종료 페이지 번호
                    }
                    else
                    {
                        xlPrinting.SAVE(vSaveFileName);
                    }
                    vMessageText = string.Format("Printing End [Total Page : {0}]", vPageNumber);
                    isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                }
                //-------------------------------------------------------------------------------------
                xlPrinting.Dispose();
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                xlPrinting.Dispose();

                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
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
                    if (IDA_RESIDENT_INCOME_ETC.IsFocused)
                    {
                        IDA_RESIDENT_INCOME_ETC.AddOver(); 
                        IGR_RESIDENT_INCOME.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_RESIDENT_INCOME_ETC.IsFocused)
                    {
                        IDA_RESIDENT_INCOME_ETC.AddUnder(); 
                        IGR_RESIDENT_INCOME.Focus();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {                    
                    IDA_RESIDENT_ETC.Update();
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_RESIDENT_INCOME_ETC.IsFocused)
                    {
                        IDA_RESIDENT_INCOME_ETC.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (IDA_RESIDENT_BUSINESS.IsFocused)
                    //{
                    //    if (IGR_RESIDENT_BSN_FAMILY.RowCount > 0)
                    //    {
                    //        IDA_RESIDENT_BSN_FAMILY.MoveFirst(IGR_RESIDENT_BSN_FAMILY.Name);
                    //        for (int C = 0; C < IGR_RESIDENT_BSN_FAMILY.RowCount; C++)
                    //        {
                    //            IDA_RESIDENT_BSN_FAMILY.Delete();
                    //            IDA_RESIDENT_BSN_FAMILY.MoveNext(IGR_RESIDENT_BSN_FAMILY.Name);
                    //        }
                    //    }
                    //    IDA_RESIDENT_BUSINESS.Delete();
                    //}
                    //else 
                    if (IDA_RESIDENT_INCOME_ETC.IsFocused)
                    {
                        IDA_RESIDENT_INCOME_ETC.Delete();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    Print_Doc("PRINT");
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    Print_Doc("FILE");
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0775_Load(object sender, EventArgs e)
        {
            IDA_RESIDENT_ETC.FillSchema();
            IDA_RESIDENT_INCOME_ETC.FillSchema();
        }

        private void HRMF0775_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2010-01");
            STD_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
        }

        private void IGR_INCOME_RESIDENT_BSN_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            if (e.ColIndex == IGR_RESIDENT_INCOME.GetColumnToIndex("PAY_DATE"))
            {
                IGR_RESIDENT_INCOME.SetCellValue("RECEIPT_DATE", e.CellValue);
            }
            else if (e.ColIndex == IGR_RESIDENT_INCOME.GetColumnToIndex("PAYMENT_AMOUNT"))
            {
                Set_Tot_Payment_Amount(iConv.ISDecimaltoZero(e.CellValue), 
                                        iConv.ISDecimaltoZero(IGR_RESIDENT_INCOME.GetCellValue("PAYMENT_ETC_AMOUNT")));
            }
            else if (e.ColIndex == IGR_RESIDENT_INCOME.GetColumnToIndex("PAYMENT_ETC_AMOUNT"))
            {
                Set_Tot_Payment_Amount(iConv.ISDecimaltoZero(IGR_RESIDENT_INCOME.GetCellValue("PAYMENT_AMOUNT")), 
                                        iConv.ISDecimaltoZero(e.CellValue));
            }
            else if (e.ColIndex == IGR_RESIDENT_INCOME.GetColumnToIndex("EXP_RATE"))
            {
                Set_Exp_Amount(iConv.ISDecimaltoZero(e.CellValue));
            }
            else if (e.ColIndex == IGR_RESIDENT_INCOME.GetColumnToIndex("INCOME_ETC_AMOUNT"))
            {
                Set_Tot_Income_Amount(iConv.ISDecimaltoZero(IGR_RESIDENT_INCOME.GetCellValue("INCOME_AMOUNT")),
                                        iConv.ISDecimaltoZero(e.CellValue));
            }
            else if (e.ColIndex == IGR_RESIDENT_INCOME.GetColumnToIndex("TAX_RATE"))
            {
                Set_Income_Tax_Amount(iConv.ISDecimaltoZero(e.CellValue));
            }
            else if (e.ColIndex == IGR_RESIDENT_INCOME.GetColumnToIndex("INCOME_TAX_AMT"))
            {
                Set_Local_Tax_Amount(iConv.ISDecimaltoZero(e.CellValue));
            }
            else if (e.ColIndex == IGR_RESIDENT_INCOME.GetColumnToIndex("LOCAL_TAX_AMT"))
            {
                Set_Tot_Dedction_Amount(iConv.ISDecimaltoZero(IGR_RESIDENT_INCOME.GetCellValue("INCOME_TAX_AMT")), 
                                        iConv.ISDecimaltoZero(e.CellValue));
            }
        }

        #endregion

        #region ----- Lookup event -----

        private void ilaNATIONALITY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("NATIONALITY_TYPE", "Y");
        }

        private void ilaINCOME_CLASS_ETC_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("INCOME_CLASS_ETC", "Y");
        }

        private void ilaBUSINESS_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BUSINESS_CODE", "Y");
        }

        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BANK", "Y");
        }

        private void ilaRELATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("RELATION", "Y");
        }

        private void ilaADDRESS_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildADDRESS.SetLookupParamValue("W_ADDRESS", ZIP_CODE.EditValue);
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        //private void IDA_RESIDENT_BUSINESS_PreRowUpdate(ISPreRowUpdateEventArgs e)
        //{
        //    if (iConv.ISNull(e.Row["NAME"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iConv.ISNull(e.Row["REPRE_NUM"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(REPRE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iConv.ISNull(e.Row["NATIONALITY_TYPE"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(NATIONALITY_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iConv.ISNull(e.Row["BUSINESS_CODE"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(BUSINESS_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //    if (iConv.ISNull(e.Row["ADDRESS1"]) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(ZIP_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        e.Cancel = true;
        //        return;
        //    }
        //}

        private void IDA_INCOME_RESIDENT_BSN_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(STD_YYYYMM_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(STD_YYYYMM_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
            if (iConv.ISNull(e.Row["PAY_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10445"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["RECEIPT_DATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10446"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["PAYMENT_AMOUNT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10447"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

        
    }
}