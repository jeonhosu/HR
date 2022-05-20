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

namespace HRMF0501
{
    public partial class HRMF0501 : Office2007Form
    {
        
        #region ----- Variables -----

        ISCommonUtil.ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();
        Object mSESSION_ID;

        #endregion;

        #region ----- Constructor -----

        public HRMF0501(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods -----

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
            ILD_CORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ILD_CORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            W_CORP_NAME.BringToFront();
        }

        private void Search_DB()
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }

            if (iConv.ISNull(W_STD_YYYYMM.EditValue) == String.Empty)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }
            IDA_GRADE_HEADER.Fill();
            IGR_GRADE_HEADER.Focus();
        }

        private void Grade_Header_Insert()
        {            
            IGR_GRADE_HEADER.SetCellValue("START_YYYYMM", W_STD_YYYYMM.EditValue);
            IGR_GRADE_HEADER.SetCellValue("ENABLED_FLAG", "Y");
        }

        private void Grade_Line_Insert()
        {
            IGR_GRADE_LINE.SetCellValue("ENABLED_FLAG", "Y");
        }
        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----        

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
                    if (IDA_GRADE_HEADER.IsFocused)
                    {
                        IDA_GRADE_HEADER.AddOver();
                        Grade_Header_Insert();     // 헤더 INSERT시 필요한값 INSERT.
                    }
                    else if (IDA_GRADE_STEP.IsFocused)
                    {
                        IDA_GRADE_STEP.AddOver();
                    }
                    else if (IDA_GRADE_LINE.IsFocused)
                    {
                        IDA_GRADE_LINE.AddOver();
                        Grade_Line_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_GRADE_HEADER.IsFocused)
                    {
                        IDA_GRADE_HEADER.AddUnder();
                        Grade_Header_Insert();     // 헤더 INSERT시 필요한값 INSERT.
                    }
                    else if (IDA_GRADE_STEP.IsFocused)
                    {
                        IDA_GRADE_STEP.AddUnder();
                    }
                    else if (IDA_GRADE_LINE.IsFocused)
                    {
                        IDA_GRADE_LINE.AddUnder();
                        Grade_Line_Insert();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_GRADE_HEADER.Update();                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_GRADE_HEADER.IsFocused)
                    {
                        IDA_GRADE_LINE.Cancel();
                        IDA_GRADE_STEP.Cancel();
                        IDA_GRADE_HEADER.Cancel();
                    }
                    else if (IDA_GRADE_STEP.IsFocused)
                    {
                        IDA_GRADE_LINE.Cancel();
                        IDA_GRADE_STEP.Cancel();
                    }
                    else if (IDA_GRADE_LINE.IsFocused)
                    {
                        IDA_GRADE_LINE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_GRADE_HEADER.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            return;
                        
                        IDC_DELETE_GRADE_HEADER.SetCommandParamValue("W_GRADE_HEADER_ID", IGR_GRADE_HEADER.GetCellValue("GRADE_HEADER_ID"));
                        IDC_DELETE_GRADE_HEADER.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_GRADE_HEADER.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_GRADE_HEADER.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_DELETE_GRADE_HEADER.ExcuteError)
                        {
                            MessageBoxAdv.Show(IDC_DELETE_GRADE_HEADER.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else if (vSTATUS.Equals("F"))
                        {
                            if (vMESSAGE != string.Empty)
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }                        
                    }
                    else if (IDA_GRADE_STEP.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            return;

                        IDC_DELETE_GRADE_STEP.SetCommandParamValue("W_GRADE_HEADER_ID", IGR_GRADE_STEP.GetCellValue("GRADE_HEADER_ID"));
                        IDC_DELETE_GRADE_STEP.SetCommandParamValue("W_GRADE_STEP", IGR_GRADE_STEP.GetCellValue("GRADE_STEP"));
                        IDC_DELETE_GRADE_STEP.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_GRADE_STEP.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_GRADE_STEP.GetCommandParamValue("O_MESSAGE"));
                        if(IDC_DELETE_GRADE_STEP.ExcuteError)
                        {
                            MessageBoxAdv.Show(IDC_DELETE_GRADE_STEP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else if(vSTATUS.Equals("F"))
                        {
                            if(vMESSAGE != string.Empty)                            
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }  
                    }
                    else if (IDA_GRADE_LINE.IsFocused)
                    {
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10030"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                            return;
                        
                        IDC_DELETE_GRADE_LINE.SetCommandParamValue("W_GRADE_LINE_ID", IGR_GRADE_LINE.GetCellValue("GRADE_LINE_ID")); 
                        IDC_DELETE_GRADE_LINE.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_DELETE_GRADE_LINE.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_DELETE_GRADE_LINE.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_DELETE_GRADE_LINE.ExcuteError)
                        {
                            MessageBoxAdv.Show(IDC_DELETE_GRADE_LINE.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else if (vSTATUS.Equals("F"))
                        {
                            if (vMESSAGE != string.Empty)
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    Search_DB();
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0501_Load(object sender, EventArgs e)
        {     
            W_STD_YYYYMM.EditValue = iDate.ISYearMonth(DateTime.Today);                        
            DefaultCorporation();              //Default Corp.
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]        

            IDC_GET_SESSION_ID_P.ExecuteNonQuery();
            mSESSION_ID = IDC_GET_SESSION_ID_P.GetCommandParamValue("O_SESSION_ID");

            IDA_GRADE_HEADER.FillSchema();
        }
        #endregion  

        #region ----- Adapter Event -----

        private void idaGRADE_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["START_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Start Year Month(시작년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PAY_TYPE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Type(급여제)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PAY_GRADE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Grade(직급)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaGRADE_HEADER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaGRADE_STEP_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["GRADE_STEP"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Grade Step(호봉)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaGRADE_STEP_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaGRADE_LINE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["ALLOWANCE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance(항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ALLOWANCE_AMOUNT"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance Amount(금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaGRADE_LINE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
      
        #endregion

        #region ----- LookUp Event -----

        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_YYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
            ILD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(DateTime.Today, 3, 0));
        }

        private void ilaPAY_GRADE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        { 
            ILD_PAY_GRADE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaPAY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPAY_GRADE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PAY_GRADE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaALLOWANCE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON_W.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE");
            ILD_COMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE9 = 'Y' ");
            ILD_COMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        private void BTN_EXCEL_EXPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0501_EXPORT vHRMF0501_EXPORT = new HRMF0501_EXPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                , W_CORP_ID.EditValue, W_CORP_NAME.EditValue
                                                                , W_STD_YYYYMM.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0501_EXPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0501_EXPORT.ShowDialog();
            vHRMF0501_EXPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void BTN_EXCEL_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0501_IMPORT vHRMF0501_IMPORT = new HRMF0501_IMPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface, W_CORP_ID.EditValue
                                                                , W_STD_YYYYMM.EditValue, mSESSION_ID);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0501_IMPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0501_IMPORT.ShowDialog();
            vHRMF0501_IMPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }
    }
}