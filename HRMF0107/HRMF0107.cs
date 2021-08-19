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

namespace HRMF0107
{
    public partial class HRMF0107 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----
        

        #endregion;

        #region ----- Constructor -----

        public HRMF0107()
        {
            InitializeComponent();
        }

        public HRMF0107(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            //-----------------------------------------------------------------------------------------------------
            // 조회 전, 조건 체크하는 부분.
            //-----------------------------------------------------------------------------------------------------
            if (iString.ISNull(STD_YEAR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YEAR_0.Focus();
                return;
            }
            //-----------------------------------------------------------------------------------------------------

            // 조건에 맞는 연말정산기준관리 데이터 출력.
            IDA_INCOME_TAX_STANDARD.Fill();
            payDefaultSetting();
                    
        }

        private void payDeleteDefaultSetting()
        {
            // [기본설정Tab]
            // 근로소득공제 항목의 Default 값 삭제.
            PAY_1.EditValue = null;
            PAY_2.EditValue = null;
            PAY_3.EditValue = null;
            PAY_4.EditValue = null;
            PAY_5.EditValue = null;
        }

        private void payDefaultSetting()
        {
            // [기본설정Tab]
            // 근로소득공제 항목의 Default 값 설정.
            if (iString.ISNull(INCOME_DED_A.EditValue) != string.Empty)
            {
                PAY_1.EditValue = 0;
            }
            else
            {
                PAY_1.EditValue = null;
            }

            if (iString.ISNull(INCOME_DED_B.EditValue) != string.Empty)
            {
                PAY_2.EditValue = Convert.ToInt32(INCOME_DED_A.EditValue) + 1;
            }
            else
            {
                PAY_2.EditValue = null;
            }

            if (iString.ISNull(INCOME_DED_C.EditValue) != string.Empty)
            {
                PAY_3.EditValue = Convert.ToInt32(INCOME_DED_B.EditValue) + 1;
            }
            else
            {
                PAY_3.EditValue = null;
            }

            if (iString.ISNull(INCOME_DED_D.EditValue) != string.Empty)
            {
                PAY_4.EditValue = Convert.ToInt32(INCOME_DED_C.EditValue) + 1;
            }
            else
            {
                PAY_4.EditValue = null;
            }

            if (iString.ISNull(INCOME_DED_LMT_0.EditValue) != string.Empty)
            {
                PAY_5.EditValue = 99999999999;
            }
            else
            {
                PAY_5.EditValue = null;
            }
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    // 데이터 등록-AddOver
                    if (IDA_INCOME_TAX_STANDARD.IsFocused)
                    {
                        IDA_INCOME_TAX_STANDARD.AddOver();
                        payDeleteDefaultSetting();
                    }                                        
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    // 데이터 등록-AddUnder
                    if (IDA_INCOME_TAX_STANDARD.IsFocused)
                    {
                        IDA_INCOME_TAX_STANDARD.AddUnder();
                        payDeleteDefaultSetting();
                    }                                        
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_INCOME_TAX_STANDARD.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_INCOME_TAX_STANDARD.IsFocused)
                    {
                        IDA_INCOME_TAX_STANDARD.Cancel();
                        payDefaultSetting();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_INCOME_TAX_STANDARD.IsFocused)
                    {
                        IDA_INCOME_TAX_STANDARD.Delete();
                        payDeleteDefaultSetting();
                    }                                                           
                }
            }
        }

        #endregion;

        //기존에 검색된 값을 제거하고 초기화 상태로 셋팅
        /*
        private void allTextBoxValueDeleteSetting()
        {
            CODE_NAME_1.EditValue = null;
            GROUP_CODE_1.EditValue = null;
            STD_YEAR_1.EditValue = null;
            DRIVE_DED_LMT_0.EditValue = null;
            RESE_DED_LMT_0.EditValue = null;
            REPORTER_DED_AMT_0.EditValue = null;
            FOREIGN_INCOME_RATE_0.EditValue = null;
            MONTH_PAY_STD_0.EditValue = null;
            OT_DED_LMT_0.EditValue = null;
            FOOD_DED_LMT_0.EditValue = null;
            BABY_DED_LMT_0.EditValue = null;
            INCOME_DED_A_0.EditValue = null;
            INCOME_CALCU_BAS_A_0.EditValue = null;
            INCOME_DED_AMT_A_0.EditValue = null;
            INCOME_DED_RATE_A_0.EditValue = null;
            INCOME_DED_B_0.EditValue = null;
            INCOME_CALCU_BAS_B_0.EditValue = null;
            INCOME_DED_AMT_B_0.EditValue = null;
            INCOME_DED_RATE_B_0.EditValue = null;
            INCOME_DED_C_0.EditValue = null;
            INCOME_CALCU_BAS_C_0.EditValue = null;
            INCOME_DED_AMT_C_0.EditValue = null;
            INCOME_DED_RATE_C_0.EditValue = null;
            INCOME_DED_D_0.EditValue = null;
            INCOME_CALCU_BAS_D_0.EditValue = null;
            INCOME_DED_AMT_D_0.EditValue = null;
            INCOME_DED_RATE_D_0.EditValue = null;
            INCOME_DED_LMT_0.EditValue = null;
            INCOME_CALCU_BAS_LMT_0.EditValue = null;
            INCOME_DED_AMT_LMT_0.EditValue = null;
            INCOME_DED_RATE_LMT_0.EditValue = null;
            PERSON_DED_AMT_0.EditValue = null;
            SPOUSE_DED_AMT_0.EditValue = null;
            SUPPORT_DED_AMT_0.EditValue = null;
            OLD_AGED_DED_AMT_0.EditValue = null;
            OLD_AGED_DED1_AMT_0.EditValue = null;
            DEFORM_DED_AMT_0.EditValue = null;
            WOMAN_DED_AMT_0.EditValue = null;
            BRING_CHILD_DED_AMT_0.EditValue = null;
            BIRTH_DED_AMT_0.EditValue = null;
            MANY_CHILD_DED_CNT_0.EditValue = null;
            MANY_CHILD_DED_BAS_AMT_0.EditValue = null;
            MANY_CHILD_DED_ADD_AMT_0.EditValue = null;
            ANCESTOR_MAN_AGE_0.EditValue = null;
            DESCENDANT_MAN_AGE_0.EditValue = null;
            OLD_DED_AGE_0.EditValue = null;
            ANCESTOR_WOMAN_AGE_0.EditValue = null;
            DESCENDANT_WOMAN_AGE_0.EditValue = null;
            OLD_DED_AGE1_0.EditValue = null;
            CHILDREN_DED_AGE_0.EditValue = null;
            BIRTH_DED_AGE_0.EditValue = null;
            MANY_CHILD_DED_AGE_0.EditValue = null;
            ETC_INSUR_LMT_0.EditValue = null;
            DEFORM_INSUR_LMT_0.EditValue = null;
            MEDIC_DED_STD_0.EditValue = null;
            MEDIC_DED_LMT_0.EditValue = null;
            PER_EDU_0.EditValue = null;
            KIND_EDU_0.EditValue = null;
            STUD_EDU_0.EditValue = null;
            UNIV_EDU_0.EditValue = null;
            HOUSE_AMT_RATE_0.EditValue = null;
            HOUSE_INTER_RATE_0.EditValue = null;
            HOUSE_AMT_LMT_0.EditValue = null;
            LONG_HOUSE_PROF_LMT_0.EditValue = null;
            LONG_HOUSE_PROF_LMT_1_0.EditValue = null;
            LONG_HOUSE_PROF_LMT_2_0.EditValue = null;
            HOUSE_TOTAL_LMT_0.EditValue = null;
            HOUSE_TOTAL_LMT_1_0.EditValue = null;
            HOUSE_TOTAL_LMT_2_0.EditValue = null;
            ASS_GIFT_RATE3_0.EditValue = null;
            ASS_GIFT_RATE3_1_0.EditValue = null;
            ASS_GIFT_RATE3_2_0.EditValue = null;
            SP_DED_STD_0.EditValue = null;
            LEGAL_GIFT_RATE_0.EditValue = null;
            ASS_GIFT_RATE1_0.EditValue = null;
            SP_DED_AMT_0.EditValue = null;
            HOUSE_MONTHLY_RATE_0.EditValue = null;
            LOW_HOUSE_ADD_RATE_0.EditValue = null;
            PRIV_PENS_RATE_0.EditValue = null;
            PRIV_PENS_LMT_0.EditValue = null;
            PENS_DED_RATE_0.EditValue = null;
            PENS_DED_LMT_0.EditValue = null;
            RETR_PENS_LMT_0.EditValue = null;
            INVEST_RATE1_0.EditValue = null;
            INVEST_LMT_RATE1_0.EditValue = null;
            INVEST_RATE2_0.EditValue = null;
            INVEST_LMT_RATE2_0.EditValue = null;
            CARD_BAS_RATE_0.EditValue = null;
            CARD_DED_RATE_0.EditValue = null;
            CHECK_CARD_DED_RATE_0.EditValue = null;
            CARD_DED_LMT_0.EditValue = null;
            CARD_DED_LMT_RATE_0.EditValue = null;
            STOCK_LMT_0.EditValue = null;
            SMALL_CORPOR_DED_LMT_0.EditValue = null;
            LONG_STOCK_SAVING_RATE_1_0.EditValue = null;
            LONG_STOCK_SAVING_RATE_2_0.EditValue = null;
            LONG_STOCK_SAVING_RATE_3_0.EditValue = null;
            LONG_STOCK_SAVING_LMT_1_0.EditValue = null;
            LONG_STOCK_SAVING_LMT_2_0.EditValue = null;
            LONG_STOCK_SAVING_LMT_3_0.EditValue = null;
            HOUSE_APP_DEPOSIT_RATE_0.EditValue = null;
            HOUSE_APP_DEPOSIT_LMT_0.EditValue = null;
            IN_TAX_STD_A_0.EditValue = null;
            IN_TAX_RATE_A_0.EditValue = null;
            IN_TAX_STD_B_0.EditValue = null;
            IN_TAX_RATE_B_0.EditValue = null;
            IN_TAX_STD_C_0.EditValue = null;
            IN_TAX_LMT_0.EditValue = null;
            IN_TAX_BASE_C_0.EditValue = null;
            POLI_GIFT_MAX_0.EditValue = null;
            POLI_GIFT_RATE_0.EditValue = null;
            POLI_GIFT_RATE1_0.EditValue = null;
            TAX_ASSO_RATE_0.EditValue = null;
            HOUSE_DEBT_BEN_RATE_0.EditValue = null;
            SP_TAX_RATE_0.EditValue = null;
            FOREIGN_TAX_DED_0.EditValue = null;
            LOCAL_TAX_RATE_0.EditValue = null;            
        }
        */

        #region ----- Form Event -----

        private void HRMF0107_Load(object sender, EventArgs e)
        {
            IDA_INCOME_TAX_STANDARD.FillSchema();            
        }

        private void HRMF0107_Shown(object sender, EventArgs e)
        {
            // Standard Year Default SETTING
            ildYEAR.SetLookupParamValue("W_START_YEAR", "2001");
            ildYEAR.SetLookupParamValue("W_END_YEAR", iDate.ISYear(DateTime.Today, 2));
            STD_YEAR_0.EditValue = iDate.ISYear(DateTime.Today);
        }
        // 근로소득공제A -- Validated
        private void INCOME_DED_A_0_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(e.EditValue) != string.Empty)
            {
                PAY_1.EditValue = 0;
                PAY_2.EditValue = Convert.ToInt32(e.EditValue) + 1;
            }
            else
            {
                PAY_1.EditValue = null;
                PAY_2.EditValue = null;
            }
        }

        // 근로소득공제B -- Validated
        private void INCOME_DED_B_0_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(e.EditValue) != string.Empty)
            {
                PAY_3.EditValue = Convert.ToInt32(e.EditValue) + 1;
            }
            else
            {
                PAY_3.EditValue = null;
            }
        }

        // 근로소득공제C -- Validated
        private void INCOME_DED_C_0_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(e.EditValue) != string.Empty)
            {
                PAY_4.EditValue = Convert.ToInt32(e.EditValue) + 1;
            }
            else
            {
                PAY_4.EditValue = null;
            }
        }

        // 근로소득공제(총급여액-부터) -- Validated
        private void INCOME_DED_LMT_0_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(e.EditValue) != string.Empty)
            {
                PAY_5.EditValue = 99999999;
            }
            else
            {
                PAY_5.EditValue = null;
            }
        }

        // 전 년도 복사 기능 버튼
        private void ibtCOPY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string mPre_YYYY;
            string mSTATUS = "F";
            string mMESSAGE = String.Empty;
            DialogResult mDialogResult;

            // '국가 정보' 및 '날짜 정보' 선택 여부 체크.
            if (iString.ISNull(STD_YEAR_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10022"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YEAR_0.Focus();
                return;
            }

            // 전 년도 자료 존재 체크.
            mPre_YYYY = iString.ISNull(iString.ISNumtoZero(STD_YEAR_0.EditValue) - 1);
            idcTAX_STANDARD_CHECK_YN.SetCommandParamValue("W_YEAR_YYYY", mPre_YYYY);
            idcTAX_STANDARD_CHECK_YN.ExecuteNonQuery();            
            mMESSAGE = iString.ISNull(idcTAX_STANDARD_CHECK_YN.GetCommandParamValue("O_CHECK_YN"));

            // 기존 자료가 존재하지 않을 경우.
            if (mMESSAGE == "N".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10083"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YEAR_0.Focus();
                return;
            }

            // 당년도 자료 존재 체크.
            idcTAX_STANDARD_CHECK_YN.SetCommandParamValue("W_YEAR_YYYY", STD_YEAR_0.EditValue);
            idcTAX_STANDARD_CHECK_YN.ExecuteNonQuery();
            mMESSAGE = iString.ISNull(idcTAX_STANDARD_CHECK_YN.GetCommandParamValue("O_CHECK_YN"));

            // 기존 자료 존재.
            if (mMESSAGE == "Y".ToString())
            {
                mDialogResult = MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (mDialogResult == DialogResult.No)
                {
                    return;
                }
            }

            // Copy 시작
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();
             
            idcTAX_STANDARD_COPY.ExecuteNonQuery();
            mSTATUS = iString.ISNull(idcTAX_STANDARD_COPY.GetCommandParamValue("O_STATUS"));
            mMESSAGE = iString.ISNull(idcTAX_STANDARD_COPY.GetCommandParamValue("O_MESSAGE"));
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();

            if (idcTAX_STANDARD_COPY.ExcuteError || mSTATUS == "F")
            { 
                MessageBoxAdv.Show(mMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } 
            if (mMESSAGE != String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("SDM_10027"), "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion







    }
}