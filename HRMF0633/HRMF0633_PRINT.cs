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
    public partial class HRMF0633_PRINT : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mCorp_ID;            //업체ID
        object mAdjustment_Type;    //정산구분
        object mPerson_ID;          //사원ID
        object mDept_ID;            //부서ID
        object mPayGrade_ID;        //직급ID

        #endregion;

        #region ----- Constructor -----

        public HRMF0633_PRINT(ISAppInterface pAppInterface
                             , object pADJUSTMENT_ID    //퇴직정산ID
                             , object pCorpID           //업체ID
                             , object pAdjustment_Type  //정산구분
                             , object pPerson_ID        //사원ID
                             , object pDept_ID          //부서ID
                             , object pPayGrade_ID      //직급ID
                             )
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            ADJUSTMENT_ID.EditValue = pADJUSTMENT_ID;
            mCorp_ID = pCorpID;            //업체ID
            mAdjustment_Type = pAdjustment_Type; //정산구분
            mPerson_ID = pPerson_ID;       //사원ID
            mDept_ID = pDept_ID;           //부서ID
            mPayGrade_ID = pPayGrade_ID;   //직급ID
        }

        #endregion;

        #region ----- Private Methods ----


        #endregion;

        #region ----- XL Export Methods ----

        private void Set_Print(string pPrint_Type)
        {
            
            //-------------------------------------------------------------------------------------
            // 명세서 선택 여부 체크
            //-------------------------------------------------------------------------------------
            if (RETIRE_ADJUSTMENT_YN.CheckBoxString != "Y" && INVOICE_WITHHOLDING_TAX_YN.CheckBoxString != "Y")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("명세서를 선택해주세요."), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RETIRE_ADJUSTMENT_YN.Focus();
                return;
            }

            if (INVOICE_WITHHOLDING_TAX_YN.CheckBoxString == "Y")
            {
                // 출력 용도 구분 체크
                if (EARNER_YN.CheckBoxString != "Y" && ADDRESSOR1_YN.CheckBoxString != "Y" && ADDRESSOR2_YN.CheckBoxString != "Y")
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("출력 용도를 한 개 이상 선택해주세요."), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    EARNER_YN.Focus();
                    return;
                }
            }

            idaRETIRE_LIST.Fill();
            
            idaRETIRE_WITHHOLDING_TAX.Fill();
            idaETC_ALLOWANCE.Fill();
            idaPRINT_2013.Fill();
            IDA_RETIRE_PRINT1.Fill();


            if (RETIRE_ADJUSTMENT_YN.CheckBoxString == "Y")
            {
                XLPrinting1(pPrint_Type);
            }

            if (INVOICE_WITHHOLDING_TAX_YN.CheckBoxString == "Y")
            {
                XLPrinting2(pPrint_Type);
            }

            //MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10035"), "", MessageBoxButtons.OK, MessageBoxIcon.None);
            // 인쇄 완료 메시지 출력     
        }

        #endregion;

        // 인쇄 부분
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
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            return vString;
        }

        #endregion;

        #region ----- XL Print Methods ----

        private void XLPrinting1(string pPrint_Type)
        {
            System.DateTime vStartTime = DateTime.Now;
            string vMessageText = string.Empty;

            int vCountRow = girdRETIRE_ADJUSTMENT.RowCount; //girdRETIRE_ADJUSTMENT 그리드의 총 행수
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            //파일 저장시 파일명 지정.
            string vSaveFileName = string.Empty;
            if (pPrint_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("RetireAdjust_");

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.Filter = "Excel file(*.xlsx)|*.xlsx";
                saveFileDialog1.DefaultExt = "xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
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


            IDC_GET_REPORT_SET.SetCommandParamValue("P_ASSEMBLY_ID", "HRMF0633");
            IDC_GET_REPORT_SET.ExecuteNonQuery();
            string vREPORT_TYPE = iString.ISNull(IDC_GET_REPORT_SET.GetCommandParamValue("O_REPORT_TYPE"));
            string vREPORT_FILE_NAME = iString.ISNull(IDC_GET_REPORT_SET.GetCommandParamValue("O_REPORT_FILE_NAME"));


            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            //iedPRINT_DATE.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting1 xlPrinting = new XLPrinting1(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                //-------------------------------------------------------------------------------------
                if(vREPORT_TYPE.ToUpper() == "NFK")
                {
                    xlPrinting.OpenFileNameExcel = vREPORT_FILE_NAME;

                }
                else
                {
                    xlPrinting.OpenFileNameExcel = "HRMF0633_001.xlsx";
                }


                //-------------------------------------------------------------------------------------

                //-------------------------------------------------------------------------------------
                // 명세서 선택 시, 해당 명세서 양식에 맞는 Printing... (다중 선택 가능)
                //-------------------------------------------------------------------------------------
                // 퇴직금 정산 내역
                if (RETIRE_ADJUSTMENT_YN.CheckBoxString == "Y")
                {
                    if(vREPORT_TYPE.ToUpper() == "NFK")
                    {
                        vPageNumber = xlPrinting.WriteRetireAdjustment_NFK(pPrint_Type, vSaveFileName, girdRETIRE_ADJUSTMENT, gridPRINT_ALLOWANCE);
                    }
                    else
                    {
                        vPageNumber = xlPrinting.WriteRetireAdjustment(pPrint_Type, vSaveFileName, girdRETIRE_ADJUSTMENT, gridETC_ALLOWANCE);
                    }
                    
                    //vPageNumber = 0; 
                }

                // 퇴직소득원천징수영수증/지급조서
                if (INVOICE_WITHHOLDING_TAX_YN.CheckBoxString == "Y")
                {
                    // 출력 용도 구분 체크
                    if (EARNER_YN.CheckBoxString != "Y" && ADDRESSOR1_YN.CheckBoxString != "Y" && ADDRESSOR2_YN.CheckBoxString != "Y")
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("출력 용도를 한 개 이상 선택해주세요."), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        EARNER_YN.Focus();
                        return;
                    }
                }     
            }
            catch (System.Exception ex)
            {
                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            //-------------------------------------------------------------------------------------
            xlPrinting.Dispose();
            //-------------------------------------------------------------------------------------

            System.DateTime vEndTime = DateTime.Now;
            System.TimeSpan vTimeSpan = vEndTime - vStartTime;

            vMessageText = string.Format("Printing End [Total Page : {0}] ---> {1}", vPageNumber, vTimeSpan.ToString());
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();


        }

        private void XLPrinting2(string pPrint_Type)
        {
            System.DateTime vStartTime = DateTime.Now;
            string vMessageText = string.Empty;

            int vCountRow = gridWITHHOLDING_TAX.RowCount; //girdRETIRE_ADJUSTMENT 그리드의 총 행수
            int vCountRow2 = gridPRINT_2013.RowCount;
            string vRetire_Year = gridPRINT_2013.GetCellValue("FINAL_RETIRE_DATE").ToString();

            if (iString.ISNumtoZero(vRetire_Year) < 2013)
            {
                if (vCountRow < 1)
                {
                    vMessageText = string.Format("Without Data");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }                 
            }
            else
            {
                if (vCountRow2 < 1)
                {
                    vMessageText = string.Format("Without Data");
                    isAppInterfaceAdv1.OnAppMessage(vMessageText);
                    System.Windows.Forms.Application.DoEvents();
                    return;
                }
            }


            //파일 저장시 파일명 지정.
            string vSaveFileName = string.Empty;
            if (pPrint_Type == "FILE")
            {
                System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                vSaveFileName = string.Format("RetireAdjust_{0}", DateTime.Today.ToShortDateString());

                saveFileDialog1.Title = "Excel Save";
                saveFileDialog1.FileName = vSaveFileName;
                saveFileDialog1.Filter = "Excel file(*.xlsx)|*.xlsx";
                saveFileDialog1.DefaultExt = "xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    vSaveFileName = saveFileDialog1.FileName;
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

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            //iedPRINT_DATE.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting2 xlPrinting = new XLPrinting2(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                if (iString.ISNumtoZero(vRetire_Year) < 2013)
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "HRMF0633_002.xlsx";
                    //-------------------------------------------------------------------------------------
                }
                else if (iString.ISNumtoZero(vRetire_Year) < 2016)
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "HRMF0633_004.xlsx";
                    //-------------------------------------------------------------------------------------
                }
                else if(iString.ISNumtoZero(vRetire_Year) < 2020)
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "HRMF0633_006.xlsx";
                    //-------------------------------------------------------------------------------------

                }
                else
                {
                    //-------------------------------------------------------------------------------------
                    xlPrinting.OpenFileNameExcel = "HRMF0633_007.xlsx";
                    //-------------------------------------------------------------------------------------

                }

                // 퇴직소득원천징수영수증/지급조서
                if (INVOICE_WITHHOLDING_TAX_YN.CheckBoxString == "Y")
                {
                    // 출력 용도 구분 체크
                    if (EARNER_YN.CheckBoxString != "Y" && ADDRESSOR1_YN.CheckBoxString != "Y" && ADDRESSOR2_YN.CheckBoxString != "Y")
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("출력 용도를 한 개 이상 선택해주세요."), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        EARNER_YN.Focus();
                        return;
                    }

                    string vPrintType = null;
                    string vPrintType_Desc = null;
                    if (EARNER_YN.CheckBoxString == "Y")
                    {
                        vPrintType = "1";
                        vPrintType_Desc = "소득자 보관용";
                        vPageNumber = xlPrinting.WriteWithholdingTax(pPrint_Type, vSaveFileName, gridWITHHOLDING_TAX, gridPRINT_2013, vPrintType, vPrintType_Desc);
                        vPageNumber = 0;
                    }
                    if (ADDRESSOR1_YN.CheckBoxString == "Y")
                    {
                        vPrintType = "2";
                        vPrintType_Desc = "발행자 보관용";
                        vPageNumber = xlPrinting.WriteWithholdingTax(pPrint_Type, vSaveFileName, gridWITHHOLDING_TAX, gridPRINT_2013, vPrintType, vPrintType_Desc);
                        vPageNumber = 0;
                    }
                    if (ADDRESSOR2_YN.CheckBoxString == "Y")
                    {
                        vPrintType = "3";
                        vPrintType_Desc = "발행자 보고용";
                        vPageNumber = xlPrinting.WriteWithholdingTax(pPrint_Type, vSaveFileName, gridWITHHOLDING_TAX, gridPRINT_2013, vPrintType, vPrintType_Desc);
                        vPageNumber = 0;
                    }
                }
            }
            catch (System.Exception ex)
            {
                vMessageText = ex.Message;
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
            }

            //-------------------------------------------------------------------------------------
            xlPrinting.Dispose();
            //-------------------------------------------------------------------------------------

            System.DateTime vEndTime = DateTime.Now;
            System.TimeSpan vTimeSpan = vEndTime - vStartTime;

            vMessageText = string.Format("Printing End [Total Page : {0}] ---> {1}", vPageNumber, vTimeSpan.ToString());
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                   
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0633_PRINT_Load(object sender, EventArgs e)
        {
            CORP_ID.EditValue = mCorp_ID;                 //업체ID
            ADJUSTMENT_TYPE.EditValue = mAdjustment_Type; //정산구분ID
            PERSON_ID.EditValue = mPerson_ID;             //사원ID
            DEPT_ID.EditValue = mDept_ID;                 //부서ID
            PAY_GRADE_ID.EditValue = mPayGrade_ID;        //직급ID
        }

        // 명세서 발급 취소
        private void btnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            // EditBox 초기화
            CORP_ID.EditValue = null;         //업체ID
            ADJUSTMENT_TYPE.EditValue = null; //정산구분ID
            PERSON_ID.EditValue = null;       //사원ID
            DEPT_ID.EditValue = null;         //부서ID
            PAY_GRADE_ID.EditValue = null;    //직급ID

            // 명세서 발급 CheckBox 초기화
            RETIRE_ADJUSTMENT_YN.CheckBoxValue = "N";
            INVOICE_WITHHOLDING_TAX_YN.CheckBoxValue = "N";
            EARNER_YN.CheckBoxValue = "N";
            ADDRESSOR1_YN.CheckBoxValue = "N";
            ADDRESSOR2_YN.CheckBoxValue = "N";

            this.Close();
        }

        // 명세서 발급
        private void btnPRINT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Print("PRINT");
        }

        private void BTN_EXCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Set_Print("FILE");
        }

        #endregion
        
    }
}