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

namespace HRMF0633
{
    public partial class HRMF0633 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISCommonUtil.ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF0633(Form pMainForm, ISAppInterface pAppInterface)
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
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

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

            if (iString.ISNull(ADJUSTMENT_TYPE_0.EditValue) == String.Empty)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ADJUSTMENT_TYPE_NAME_0.Focus();
                return;
            }
            idaPERSON.Fill();
            idaRETIRE_ADJUSTMENT.Fill();
            igrPERSON.Focus();            
        }

        private void Insert_Retire_Cal()
        {    
            RETIRE_DATE_FR.EditValue = igrPERSON.GetCellValue("START_DATE");
            RETIRE_DATE_TO.EditValue = igrPERSON.GetCellValue("END_DATE");
            PAY_DATE_TO.EditValue = igrPERSON.GetCellValue("END_DATE");

            CORP_ID.EditValue = W_CORP_ID.EditValue;
            PERSON_ID.EditValue = igrPERSON.GetCellValue("PERSON_ID");
            ADJUSTMENT_TYPE.EditValue = ADJUSTMENT_TYPE_0.EditValue;
        }

        private void Insert_Pay()
        {
            igrPAYMENT_SALARY.SetCellValue("ADJUSTMENT_ID", ADJUSTMENT_DC_ID.EditValue);
            igrPAYMENT_SALARY.SetCellValue("WAGE_TYPE", WAGE_TYPE_P1.EditValue);
        }

        private void Insert_Bonus()
        {
            igrPAYMENT_ETC.SetCellValue("ADJUSTMENT_ID", ADJUSTMENT_DC_ID.EditValue);
            igrPAYMENT_ETC.SetCellValue("WAGE_TYPE", WAGE_TYPE_P2.EditValue);
        }

        private void insert_Etc_Allowance()
        {
           // igrETC_ALLOWANCE.SetCellValue("ADJUSTMENT_ID", ADJUSTMENT_ID.EditValue);
            //igrETC_ALLOWANCE.SetCellValue("PERSIN_ID", RETIRE_PERSON_ID.EditValue);
        }

        private void SET_RE_CALCULATE(object pRETIRE_CAL_TYPE)
        {
            if(iString.ISNull(RETIRE_DATE_FR.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RETIRE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(RETIRE_DATE_TO.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RETIRE_DATE_TO.Focus();
                return;
            }
            if (iString.ISNull(PAY_DATE_TO.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_DATE_TO.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();
            
            //idaPAYMENT_SALARY.Update();
            //idaPAYMENT_ETC.Update();
            idaRETIRE_ADJUSTMENT.Update();

            // 실행.
            string mStatus = "F";
            string mMessage = null;
            isDataTransaction1.BeginTran();
           // idcRETIRE_CALCULATE.SetCommandParamValue("W_RETIRE_CAL_TYPE", pRETIRE_CAL_TYPE);
            idcRETIRE_CALCULATE.ExecuteNonQuery();
            mStatus = iString.ISNull(idcRETIRE_CALCULATE.GetCommandParamValue("O_STATUS"));
            mMessage = iString.ISNull(idcRETIRE_CALCULATE.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (idcRETIRE_CALCULATE.ExcuteError || mStatus == "F")
            {
                isDataTransaction1.RollBack();
                if (mMessage != string.Empty)
                {
                    MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            isDataTransaction1.Commit();
            if (mMessage != string.Empty)
            {
                MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            idaRETIRE_ADJUSTMENT.Fill();
        }

        private void isOnPrinting()
        {
            if (W_CORP_ID.EditValue == null) // 업체명
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }

            if (iString.ISNull(ADJUSTMENT_TYPE_0.EditValue) == String.Empty) // 정산구분
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ADJUSTMENT_TYPE_NAME_0.Focus();
                return;
            }

            if (iString.ISNull(PERSON_NAME.EditValue) == String.Empty) // 직원 선택 여부 체크
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("해당 인원에 대한 정보를 선택해주세요."), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                igrPERSON.Focus();
                return;
            }

            // Child Form Load
            DialogResult vdlgResult;
            Form vHRMF0633_PRINT = new HRMF0633_PRINT(isAppInterfaceAdv1.AppInterface
                                                     , ADJUSTMENT_DC_ID.EditValue
                                                     , W_CORP_ID.EditValue
                                                     , ADJUSTMENT_TYPE_0.EditValue
                                                     , igrPERSON.GetCellValue("PERSON_ID")
                                                     , igrPERSON.GetCellValue("DEPT_ID")
                                                     , igrPERSON.GetCellValue("PAY_GRADE_ID")
                                                     );

            vdlgResult = vHRMF0633_PRINT.ShowDialog();
            if (vdlgResult == DialogResult.OK)
            { }
            vHRMF0633_PRINT.Dispose();
        }

        #endregion;

        #region ------ Initialize -----

        private void Init_Real_Amount()
        {            
            //// 실 총지급액 정리.
            //REAL_TOTAL_AMOUNT.EditValue = iString.ISDecimaltoZero(REAL_AMOUNT.EditValue) + iString.ISDecimaltoZero(H_REAL_AMOUNT.EditValue)
            //                            - iString.ISDecimaltoZero(ETC_DED_AMOUNT.EditValue);
            //if (iString.ISDecimaltoZero(REAL_TOTAL_AMOUNT.EditValue) < 0)
            //{
            //    REAL_TOTAL_AMOUNT.EditValue = 0;
            //}
        }

        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events ------

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
                    if (idaPERSON.IsFocused || idaRETIRE_ADJUSTMENT.IsFocused)
                    {
                        idaRETIRE_ADJUSTMENT.AddOver();
                        Insert_Retire_Cal();
                    }
                    //else if (idaPAYMENT_PAY.IsFocused)
                    //{
                    //    idaPAYMENT_PAY.AddOver();
                    //    Insert_Pay();
                    //}
                    //else if (idaPAYMENT_BONUS.IsFocused)
                    //{
                    //    idaPAYMENT_BONUS.AddOver();
                    //    Insert_Bonus();
                    //}
                    //else if (idaSALARY_DETAIL.IsFocused)
                    //{
                    //    idaSALARY_DETAIL.AddOver();
                    //    //Insert_Pay_Detail();
                    //}
                
                    //else if (idaETC_ALLOWANCE.IsFocused)
                    //{
                    //    idaETC_ALLOWANCE.AddOver();
                    //    insert_Etc_Allowance();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaPERSON.IsFocused || idaRETIRE_ADJUSTMENT.IsFocused)
                    {
                        idaRETIRE_ADJUSTMENT.AddUnder();
                        Insert_Retire_Cal();
                    }
                    //else if (idaPAYMENT_PAY.IsFocused)
                    //{
                    //    idaPAYMENT_PAY.AddUnder();
                    //    Insert_Pay();
                    //}
                    //else if (idaPAYMENT_BONUS.IsFocused)
                    //{
                    //    idaPAYMENT_BONUS.AddUnder();
                    //    Insert_Bonus();
                    //}
                    //else if (idaSALARY_DETAIL.IsFocused)
                    //{
                    //    idaSALARY_DETAIL.AddUnder();
                    //    //Insert_Pay_Detail();l;
                    //}
               
                    //else if (idaETC_ALLOWANCE.IsFocused)
                    //{
                    //    idaETC_ALLOWANCE.AddUnder();
                    //    insert_Etc_Allowance();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        //if (idaPERSON.IsFocused || idaRETIRE_ADJUSTMENT.IsFocused)
                        //{
                        //    idaRETIRE_ADJUSTMENT.Update();
                        //}
                        //else if (idaPAYMENT_SALARY.IsFocused)
                        //{
                        //    idaPAYMENT_SALARY.Update();
                        //}
                        //else if (idaPAYMENT_ETC.IsFocused)
                        //{
                        //    idaPAYMENT_ETC.Update();
                        //}
                        idaRETIRE_ADJUSTMENT.Update();
                        //idaSALARY_DETAIL.Update();
                        //idaETC_ALLOWANCE.Update();
                    }
                    catch (Exception Ex)
                    {
                        isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaPERSON.IsFocused || idaRETIRE_ADJUSTMENT.IsFocused)
                    {
                        idaRETIRE_ADJUSTMENT.Cancel();
                    }
                    else if (idaPAYMENT_SALARY.IsFocused)
                    {
                        idaPAYMENT_SALARY.Cancel();
                    }
                    else if (idaPAYMENT_ETC.IsFocused)
                    {
                        idaPAYMENT_ETC.Cancel();
                    }
                    else if (idaSALARY_DETAIL.IsFocused)
                    {
                        idaSALARY_DETAIL.Cancel();
                    }                
                    else if (idaETC_ALLOWANCE.IsFocused)
                    {
                        idaETC_ALLOWANCE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaPERSON.IsFocused || idaRETIRE_ADJUSTMENT.IsFocused)
                    {
                        idaRETIRE_ADJUSTMENT.Delete();
                    }
                    //else if (idaPAYMENT_SALARY.IsFocused)
                    //{
                    //    idaPAYMENT_SALARY.Delete();
                    //}
                    //else if (idaPAYMENT_ETC.IsFocused)
                    //{
                    //    idaPAYMENT_ETC.Delete();
                    //}
                    //else if (idaSALARY_DETAIL.IsFocused)
                    //{
                    //    idaSALARY_DETAIL.Delete();
                    //}
                    //else if (idaETC_ALLOWANCE.IsFocused)
                    //{
                    //    idaETC_ALLOWANCE.Delete();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    isOnPrinting();
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0633_Load(object sender, EventArgs e)
        {
            //idaPAY_MASTER_HEADER.FillSchema();
            idaPERSON.FillSchema();
            
            DefaultCorporation();              //Default Corp.
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]           
        }

        private void ETC_DED_AMOUNT_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            Init_Real_Amount();
        }

        private void ibtPAYMENT_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            idaSALARY_DETAIL.Fill();
        }

        private void ibtRETIRE_CALCULATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SET_RE_CALCULATE("NEW");
        }
        

        private void ibtCLOSED_YN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //idaADJUSTMENT_CLOSED.Update();
        }

        private void btnPAYMENT_PERIOD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //if (iString.ISNull(ADJUSTMENT_DC_ID.EditValue) == string.Empty)
            //{
            //    return;
            //}

            //DialogResult dlgResult;
            //HRMF0633_PAYMENT_PERIOD vHRMF0633_PAYMENT_PERIOD = new HRMF0633_PAYMENT_PERIOD(isAppInterfaceAdv1.AppInterface, ADJUSTMENT_DC_ID.EditValue, "P1");

            //dlgResult = vHRMF0633_PAYMENT_PERIOD.ShowDialog();
            //if (dlgResult == DialogResult.OK)
            //{
            //}
            //vHRMF0633_PAYMENT_PERIOD.Dispose();
        }
        
        private void RETIRE_DATE_TO_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            PAY_DATE_TO.EditValue = RETIRE_DATE_TO.EditValue;
        }

        #endregion  

        #region ----- LookUp Event -----

        private void ilaADJUSTMENT_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "RETIRE_ADJUSTMENT_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_YYYYMM_W_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_STD_YYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(DateTime.Today, 3)));
        }

        private void ilaPAY_GRADE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "POST");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaALLOWANCE_PAY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaALLOWANCE_BONUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaALLOWANCE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "ALLOWANCE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void idaPAYMENT_PAY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //if (iString.ISNull(e.Row["ADJUSTMENT_DC_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }
        
        private void idaPAYMENT_BONUS_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            //if (iString.ISNull(e.Row["ADJUSTMENT_DC_ID"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10023"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["PAY_YYYYMM"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10107"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            //if (iString.ISNull(e.Row["WAGE_TYPE"]) == string.Empty)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10105"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
        }
        
        private void idaPERSON_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            idaRETIRE_ADJUSTMENT.Fill();
        }

        private void idaETC_ALLOWANCE_UpdateCompleted(object pSender)
        {
            idaRETIRE_ADJUSTMENT.Fill();
        }
        

        private void idaADJUSTMENT_CLOSED_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {

        }

        #endregion

        private void igrPAYMENT_SALARY_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            switch (igrPAYMENT_SALARY.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "TOTAL_AMOUNT":  // 이 인덱스가 바뀌면
                                         // 추가, 바뀐값 x 곱할값
                    igrPAYMENT_SALARY.SetCellValue("YEAR_AVG_AMOUNT", Convert.ToDecimal(e.NewValue)/12);

                    int vSum_Salary = 0;
                    object vSalary = 0;
                    for(int i = 0 ; igrPAYMENT_SALARY.RowCount >= i  ; i++ )
                    {

                        vSalary = igrPAYMENT_SALARY.GetCellValue(i, igrPAYMENT_SALARY.GetColumnToIndex("YEAR_AVG_AMOUNT"));
                        vSum_Salary = vSum_Salary + iConv.ISNumtoZero(vSalary);
                    }
                    RETIRE_AMOUNT.EditValue = vSum_Salary;

                    RETIRE_AMOUNT_EditValueChanged();
                    TOTAL_RETIRE_AMOUNT_EditValueChanged();
                    break;

                //case "DAY_OVER_CAPA":
                //    igrPAYMENT_SALARY.SetCellValue("SUMMARY", Convert.ToDecimal(e.NewValue) + Convert.ToDecimal(IGR_RETIRE_RESERVE_DC.GetCellValue("DAY_NORMAL_CAPA")) + Convert.ToDecimal(IGR_RETIRE_RESERVE_DC.GetCellValue("NIGHT_NORMAL_CAPA")) + Convert.ToDecimal(IGR_RETIRE_RESERVE_DC.GetCellValue("NIGHT_OVER_CAPA")));

                //    break;

            
                default:
                    break;


            }
        }

        private void RETIRE_AMOUNT_EditValueChanged()
        {
            // 신청금액 변경
            object vETC_AMOUNT = ETC_AMOUNT.EditValue;  //총합계 (b)
            object vRETIRE_AMOUNTT = RETIRE_AMOUNT.EditValue; // 총합계 (a)
            int vTOTAL_AMOUNT = 0;
            vTOTAL_AMOUNT = iConv.ISNumtoZero(vRETIRE_AMOUNTT) + iConv.ISNumtoZero(vETC_AMOUNT);
            TOTAL_RETIRE_AMOUNT.EditValue = vTOTAL_AMOUNT; //신청금액 (c) = (a) + (b)
        }

        private void TOTAL_RETIRE_AMOUNT_EditValueChanged()
        {
            // 차액변경      
            object vREAL_TRANSFER = REAL_TRANSFER_AMOUNT.EditValue;  //최종불입금 (d)
            object vTOTAL_RETIRE_AMOUNT = TOTAL_RETIRE_AMOUNT.EditValue;  //신청금액 (c) = (a) + (b)
            int vREAL_RETIRE_AMT = 0;
            vREAL_RETIRE_AMT = iConv.ISNumtoZero(vTOTAL_RETIRE_AMOUNT) - iConv.ISNumtoZero(vREAL_TRANSFER);
            REAL_RETIRE_AMOUNT.EditValue = vREAL_RETIRE_AMT;  // 차액 (c) - (d) 
        }

        private void ADD_TRANSFER_AMOUNT_EditValueChanged(object pSender)
        {
            //// 추가납입 변경 시 최종불입액 변경
            //object vPRE_TRANSFER = PRE_BANK_TRANSFER_AMOUNT.EditValue;  // 불입액
            //object vADD_TRANS_AMT = ADD_TRANSFER_AMOUNT.EditValue;  //추가납입액
            //int vREAL_RETIRE_AMT = 0;
            //vREAL_RETIRE_AMT = iConv.ISNumtoZero(vADD_TRANS_AMT) + iConv.ISNumtoZero(vPRE_TRANSFER);
            //REAL_RETIRE_AMOUNT.EditValue = vREAL_RETIRE_AMT;

            //TOTAL_RETIRE_AMOUNT_EditValueChanged();

        }

        private void igrPAYMENT_SALARY_CellDoubleClick(object pSender)
        {
            if (iString.ISNull(ADJUSTMENT_DC_ID.EditValue) == string.Empty)
            {
                return;
            }
            object vSALARY_YYYY = igrPAYMENT_SALARY.GetCellValue("SALARY_YYYY");
            DialogResult dlgResult;
            HRMF0633_PAYMENT_PERIOD vHRMF0633_PAYMENT_PERIOD = new HRMF0633_PAYMENT_PERIOD(isAppInterfaceAdv1.AppInterface, ADJUSTMENT_DC_ID.EditValue, "P1",  vSALARY_YYYY);

            dlgResult = vHRMF0633_PAYMENT_PERIOD.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
            }
            vHRMF0633_PAYMENT_PERIOD.Dispose();

        }

        private void isButton1_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //지급처리 

            if (iString.ISNull(RETIRE_DATE_FR.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RETIRE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(RETIRE_DATE_TO.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RETIRE_DATE_TO.Focus();
                return;
            }
            if (iString.ISNull(PAY_DATE_TO.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_DATE_TO.Focus();
                return;
            }
            if (iString.ISNull(CLOSED_DATE.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10445"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CLOSED_DATE.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();
            
            idaRETIRE_ADJUSTMENT.Update();

            // 실행.
            string mStatus = "F";
            string mMessage = null;
            isDataTransaction1.BeginTran();
            IDC_RETIRE_CLOSED.SetCommandParamValue("P_CLOSED_FLAG", "Y");
            IDC_RETIRE_CLOSED.ExecuteNonQuery();
            mStatus = iString.ISNull(IDC_RETIRE_CLOSED.GetCommandParamValue("O_STATUS"));
            mMessage = iString.ISNull(IDC_RETIRE_CLOSED.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (IDC_RETIRE_CLOSED.ExcuteError || mStatus == "F")
            {
                isDataTransaction1.RollBack();
                if (mMessage != string.Empty)
                {
                    MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            isDataTransaction1.Commit();
            if (mMessage != string.Empty)
            {
                MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

           
            idaRETIRE_ADJUSTMENT.Fill();
        }



        private void isButton2_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //지급취소 
            if (iString.ISNull(RETIRE_DATE_FR.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RETIRE_DATE_FR.Focus();
                return;
            }
            if (iString.ISNull(RETIRE_DATE_TO.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RETIRE_DATE_TO.Focus();
                return;
            }
            if (iString.ISNull(PAY_DATE_TO.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PAY_DATE_TO.Focus();
                return;
            }
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            idaRETIRE_ADJUSTMENT.Update();

            // 실행.
            string mStatus = "F";
            string mMessage = null;
            isDataTransaction1.BeginTran();
            IDC_RETIRE_CLOSED.SetCommandParamValue("P_CLOSED_FLAG", "N");
            IDC_RETIRE_CLOSED.ExecuteNonQuery();
            mStatus = iString.ISNull(IDC_RETIRE_CLOSED.GetCommandParamValue("O_STATUS"));
            mMessage = iString.ISNull(IDC_RETIRE_CLOSED.GetCommandParamValue("O_MESSAGE"));

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            if (IDC_RETIRE_CLOSED.ExcuteError || mStatus == "F")
            {
                isDataTransaction1.RollBack();
                if (mMessage != string.Empty)
                {
                    MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            isDataTransaction1.Commit();
            if (mMessage != string.Empty)
            {
                MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


            idaRETIRE_ADJUSTMENT.Fill();
        }

        private void igrPAYMENT_ETC_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            switch (igrPAYMENT_ETC.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "ETC_AMOUNT":  // 이 인덱스가 바뀌면
                                      // 추가, 바뀐값 x 곱할값
                    igrPAYMENT_ETC.SetCellValue("ETC_AMOUNT", Convert.ToDecimal(e.NewValue));

                    int vSum_Salary = 0;
                    object vSalary = 0;
                    for (int i = 0; igrPAYMENT_ETC.RowCount >= i; i++)
                    {

                        vSalary = igrPAYMENT_ETC.GetCellValue(i, igrPAYMENT_ETC.GetColumnToIndex("ETC_AMOUNT"));
                        vSum_Salary = vSum_Salary + iConv.ISNumtoZero(vSalary);
                    }
                    ETC_AMOUNT.EditValue = vSum_Salary;

                    RETIRE_AMOUNT_EditValueChanged();
                    TOTAL_RETIRE_AMOUNT_EditValueChanged();
                    break;

                //case "DAY_OVER_CAPA":
                //    igrPAYMENT_SALARY.SetCellValue("SUMMARY", Convert.ToDecimal(e.NewValue) + Convert.ToDecimal(IGR_RETIRE_RESERVE_DC.GetCellValue("DAY_NORMAL_CAPA")) + Convert.ToDecimal(IGR_RETIRE_RESERVE_DC.GetCellValue("NIGHT_NORMAL_CAPA")) + Convert.ToDecimal(IGR_RETIRE_RESERVE_DC.GetCellValue("NIGHT_OVER_CAPA")));

                //    break;


                default:
                    break;


            }
        }

    }
}