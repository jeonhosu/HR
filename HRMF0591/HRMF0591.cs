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
using System.IO;
using Syncfusion.GridExcelConverter;
using Syncfusion.XlsIO;

namespace HRMF0591
{
    public partial class HRMF0591 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        #region ----- Constructor -----
        public HRMF0591(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;

        #region ----- Property / Method ----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            IDC_DEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            IDC_DEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            IDC_DEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = IDC_DEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront(); 
        }

        private void DefaultInfo()
        {
            //일자//
            IDC_GET_LOCAL_DATE_P.ExecuteNonQuery();
            V_EXCH_DATE.EditValue = IDC_GET_LOCAL_DATE_P.GetCommandParamValue("X_LOCAL_DATE");

            //계산통화//
            IDC_GET_CAL_CURRENCY_P.SetCommandParamValue("P_STD_DATE", V_EXCH_DATE.EditValue);
            IDC_GET_CAL_CURRENCY_P.SetCommandParamValue("P_HR_MODULE", "25");
            IDC_GET_CAL_CURRENCY_P.ExecuteNonQuery();
            V_CURRENCY_CODE.EditValue = IDC_GET_CAL_CURRENCY_P.GetCommandParamValue("O_CURRENCY_CODE");

            //환율//
            IDC_GET_EXCHANGE_RATE_DATE_P.SetCommandParamValue("P_APPLY_DATE", V_EXCH_DATE.EditValue);
            IDC_GET_EXCHANGE_RATE_DATE_P.SetCommandParamValue("P_CURRENCY_CODE", V_CURRENCY_CODE.EditValue);
            IDC_GET_EXCHANGE_RATE_DATE_P.ExecuteNonQuery();
            V_EXCHANGE_RATE.EditValue = IDC_GET_EXCHANGE_RATE_DATE_P.GetCommandParamValue("X_EXCHANGE_RATE"); 
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (W_YYYYMM.EditValue == null)
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDC_GET_LOCAL_DATE_P.ExecuteNonQuery();


            if(TB_MAIN.SelectedTab.TabIndex == TP_RENTAL_MASTER.TabIndex)
            {
                string vPERSON_NUM = iString.ISNull(IGR_HOUSE_RENTAL.GetCellValue("PERSON_NUM"));

                IGR_HOUSE_RENTAL.LastConfirmChanges();
                IDA_HOUSE_RENTAL.OraSelectData.AcceptChanges();
                IDA_HOUSE_RENTAL.Refillable = true;

                IDA_HOUSE_RENTAL.Fill();
                
                IGR_HOUSE_RENTAL.Focus();
                if (IGR_HOUSE_RENTAL.RowCount > 1)
                {
                    int vIDX_PERSON_ID = IGR_HOUSE_RENTAL.GetColumnToIndex("PERSON_NUM");
                    for (int i = 0; i < IGR_HOUSE_RENTAL.RowCount; i++)
                    {
                        if (vPERSON_NUM == iString.ISNull(IGR_HOUSE_RENTAL.GetCellValue(i, vIDX_PERSON_ID)))
                        {
                            IGR_HOUSE_RENTAL.CurrentCellMoveTo(i, IGR_HOUSE_RENTAL.GetColumnToIndex("NAME"));
                            return;
                        }
                    }
                } 
            }
            else
            { 
                if (iString.ISNull(W_WAGE_TYPE.EditValue) == string.Empty)
                {// 급상여 구분
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_WAGE_TYPE_NAME.Focus();
                    return;
                }
                string vPERSON_ID = iString.ISNull(IGR_HOUSE_RENTAL_FEE.GetCellValue("PERSON_ID"));
                IDA_HOUSE_RENTAL_FEE.Fill();
                 
                IGR_HOUSE_RENTAL_FEE.Focus();
                if (IGR_HOUSE_RENTAL_FEE.RowCount > 1)
                {
                    int vIDX_PERSON_ID = IGR_HOUSE_RENTAL_FEE.GetColumnToIndex("PERSON_ID");
                    for (int i = 0; i < IGR_HOUSE_RENTAL_FEE.RowCount; i++)
                    {
                        if (vPERSON_ID == iString.ISNull(IGR_HOUSE_RENTAL_FEE.GetCellValue(i, vIDX_PERSON_ID)))
                        {
                            IGR_HOUSE_RENTAL_FEE.CurrentCellMoveTo(i, IGR_HOUSE_RENTAL_FEE.GetColumnToIndex("NAME"));
                            return;
                        }
                    }
                }
            }



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
            return mPrompt;
        }

        #endregion;

        #region ----- Excel Export -----

        private void ExcelExport(ISGridAdvEx pGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            GridExcelConverterControl vExport = new GridExcelConverterControl();

            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.Title = "Save File Name";
            saveFileDialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Application.UseWaitCursor = true;
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();

                //xls 저장방법
                //vExport.GridToExcel(pGrid.BaseGrid, saveFileDialog.FileName,
                //                    Syncfusion.GridExcelConverter.ConverterOptions.ColumnHeaders);



                //if (MessageBox.Show("Do you wish to open the xls file now?",
                //                    "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                //    vProc.StartInfo.FileName = saveFileDialog.FileName;
                //    vProc.Start();
                //}

                //xlsx 파일 저장 방법
                GridExcelConverterControl converter = new GridExcelConverterControl();
                ExcelEngine excelEngine = new ExcelEngine();
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workBook = ExcelUtils.CreateWorkbook(1);
                workBook.Version = ExcelVersion.Excel2013;
                IWorksheet sheet = workBook.Worksheets[0];
                //used to convert grid to excel 
                converter.GridToExcel(pGrid.BaseGrid, sheet, ConverterOptions.ColumnHeaders);
                //used to save the file
                workBook.SaveAs(saveFileDialog.FileName);

                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                if (MessageBox.Show("Do you wish to open the xls file now?",
                                        "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process vProc = new System.Diagnostics.Process();
                    vProc.StartInfo.FileName = saveFileDialog.FileName;
                    vProc.Start();
                }
            }
        }

        #endregion


        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----

        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_HOUSE_RENTAL.IsFocused)
                    {
                        try
                        {
                            IDA_HOUSE_RENTAL.Update();
                        }
                        catch 
                        {
                            return;
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_HOUSE_RENTAL_FEE.IsFocused)
                    {
                        IDA_HOUSE_RENTAL_FEE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {

                }
                else if(e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    if(IDA_HOUSE_RENTAL.IsFocused)
                    {
                        ExcelExport(IGR_HOUSE_RENTAL);
                    }
                    else if(IDA_HOUSE_RENTAL_FEE.IsFocused)
                    {
                        ExcelExport(IGR_HOUSE_RENTAL_FEE);
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void HRMF0591_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0591_Shown(object sender, EventArgs e)
        {
            W_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]

            DefaultCorporation();                  // Corp Default Value Setting.
                                                   // FillSchema
            IDA_HOUSE_RENTAL.FillSchema();
            IDA_HOUSE_RENTAL_FEE.FillSchema();

            // LEAVE CLOSE TYPE SETTING
            ildAPPROVAL_STATUS_0.SetLookupParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            ildAPPROVAL_STATUS_0.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            idcDEFAULT_VALUE.SetCommandParamValue("W_GROUP_CODE", "LEAVE_CLOSE_TYPE");
            idcDEFAULT_VALUE.ExecuteNonQuery();
            W_APPROVAL_STATUS_NAME.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE_NAME").ToString();
            W_APPROVAL_STATUS.EditValue = idcDEFAULT_VALUE.GetCommandParamValue("O_CODE").ToString();

            RB_ENROLLED.CheckedState = ISUtil.Enum.CheckedState.Checked;
            V_STATUS.EditValue = RB_ENROLLED.RadioButtonString;

            DefaultInfo();
        }

        private void RB_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv iStatus = sender as ISRadioButtonAdv;
            V_STATUS.EditValue = iStatus.RadioButtonString; 
        }

        private void SET_INSURANCE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_YYYYMM.EditValue) == String.Empty)
            {// 보험료년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YYYYMM.Focus();
                return;
            }
            if (iString.ISNull(W_WAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WAGE_TYPE_NAME.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0591_SET = new HRMF0591_SET(isAppInterfaceAdv1.AppInterface, "CAL"
                                                , W_CORP_ID.EditValue, W_CORP_NAME.EditValue
                                                , W_YYYYMM.EditValue
                                                , W_WAGE_TYPE.EditValue, W_WAGE_TYPE_NAME.EditValue
                                                , W_OPERATING_UNIT_DESC.EditValue, W_OPERATING_UNIT_ID.EditValue
                                                , W_DEPT_ID.EditValue, W_DEPT_NAME.EditValue
                                                , W_PERSON_ID.EditValue, W_PERSON_NUM.EditValue, W_PERSON_NAME.EditValue
                                                , V_EXCH_DATE.EditValue, V_EXCHANGE_RATE.EditValue, V_CURRENCY_CODE.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0591_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0591_SET.ShowDialog();
            vHRMF0591_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void V_EXCH_DATE_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            IDC_GET_EXCHANGE_RATE_DATE_P.ExecuteNonQuery();
            V_EXCHANGE_RATE.EditValue = IDC_GET_EXCHANGE_RATE_DATE_P.GetCommandParamValue("X_EXCHANGE_RATE");
        }

        private void ibtSET_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YYYYMM.Focus();
                return;
            }
            if (iString.ISNull(W_WAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WAGE_TYPE_NAME.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0591_SET = new HRMF0591_SET(isAppInterfaceAdv1.AppInterface, "CLOSE"
                                                , W_CORP_ID.EditValue, W_CORP_NAME.EditValue
                                                , W_YYYYMM.EditValue
                                                , W_WAGE_TYPE.EditValue, W_WAGE_TYPE_NAME.EditValue
                                                , W_OPERATING_UNIT_DESC.EditValue, W_OPERATING_UNIT_ID.EditValue
                                                , W_DEPT_ID.EditValue, W_DEPT_NAME.EditValue
                                                , W_PERSON_ID.EditValue, W_PERSON_NUM.EditValue, W_PERSON_NAME.EditValue
                                                , V_EXCH_DATE.EditValue, V_EXCHANGE_RATE.EditValue, V_CURRENCY_CODE.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0591_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0591_SET.ShowDialog();
            vHRMF0591_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void ibtSET_CANCEL_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_YYYYMM.EditValue) == String.Empty)
            {// 급여년월
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_YYYYMM.Focus();
                return;
            }
            if (iString.ISNull(W_WAGE_TYPE.EditValue) == string.Empty)
            {// 급상여 구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_WAGE_TYPE_NAME.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult vdlgResult;
            Form vHRMF0591_SET = new HRMF0591_SET(isAppInterfaceAdv1.AppInterface, "CLOSED_CANCEL"
                                                , W_CORP_ID.EditValue, W_CORP_NAME.EditValue
                                                , W_YYYYMM.EditValue
                                                , W_WAGE_TYPE.EditValue, W_WAGE_TYPE_NAME.EditValue
                                                , W_OPERATING_UNIT_DESC.EditValue, W_OPERATING_UNIT_ID.EditValue
                                                , W_DEPT_ID.EditValue, W_DEPT_NAME.EditValue
                                                , W_PERSON_ID.EditValue, W_PERSON_NUM.EditValue, W_PERSON_NAME.EditValue
                                                , V_EXCH_DATE.EditValue, V_EXCHANGE_RATE.EditValue, V_CURRENCY_CODE.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0591_SET, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0591_SET.ShowDialog();
            vHRMF0591_SET.Dispose();

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        #endregion

        #region ----- Data Adapter Event -----

        private void IDA_INSUR_AMOUNT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_ID"]) == string.Empty)
            {// 사원.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["INSUR_YYYYMM"]) == string.Empty)
            {//
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            {// cc
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(W_WAGE_TYPE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void IDA_INSUR_AMOUNT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
         
        #endregion

        #region ----- Lookup Event -----

        private void ILA_DEPT_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        { 
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaWAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CLOSING_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 'PAY' ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_W_OPERATING_UNIT_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        #endregion

    }
}