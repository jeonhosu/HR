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

namespace HRMF0705
{
    public partial class HRMF0705_PRINT : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        private string mRadioValue = string.Empty;

        int mCorp_ID;           //업체ID
        string mPrint_Num;      //발급번호
        DateTime mPrint_Date;   //발급일자
        string mPrint_Year;     //징수년도
        string mCert_Type_Name; //증명서구분
        object mCert_Type_ID;   //증명서 구분ID
        object mCert_Type_Code; //증명서 코드
        string mName;           //성명
        object mPerson_ID;      //사원ID
        DateTime mJoin_Date;    //입사일자
        DateTime mRetire_Date;  //퇴사일자
        string mDescription;    //용도
        string mSend_Org;       //제출처
        int mPrint_Count;       //매수
        string mOutChoice;      //출력방향[프린터/파일]

        #endregion;

        #region ----- Constructor -----

        public HRMF0705_PRINT(ISAppInterface pAppInterface,
                              int pCorp_ID,           //업체ID
                              string pPrint_Num,      //발급번호
                              DateTime pPrint_Date,   //발급일자
                              string pPrint_Year,     //징수년도
                              string pCert_Type_Name, //증명서구분
                              object pCert_Type_ID,   //증명서 구분ID
                              object pCert_Type_Code, //증명서 코드
                              string pName,           //성명
                              object pPerson_ID,      //사원ID
                              DateTime pJoin_Date,    //입사일자
                              DateTime pRetire_Date,  //퇴사일자
                              string pDescription,    //용도
                              string pSend_Org,       //제출처
                              int pPrint_Count,       //매수
                              string pOutChoice       //출력방향[프린터/파일]
                             )
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            mCorp_ID = pCorp_ID;               //업체ID
            mPrint_Num = pPrint_Num;           //발급번호
            mPrint_Date = pPrint_Date;         //발급일자
            mPrint_Year = pPrint_Year;         //징수년도
            mCert_Type_Name = pCert_Type_Name; //증명서구분
            mCert_Type_ID = pCert_Type_ID;     //증명서 구분ID
            mCert_Type_Code = pCert_Type_Code; //증명서 코드
            mName = pName;                     //성명

            if (Convert.ToInt32(pPerson_ID) == 0)
            {
                mPerson_ID = null;             //사원ID
            }
            else
            {
                mPerson_ID = pPerson_ID;       //사원ID
            }

            mJoin_Date = pJoin_Date;           //입사일자
            mRetire_Date = pRetire_Date;       //퇴사일자
            mDescription = pDescription;       //용도
            mSend_Org = pSend_Org;             //제출처
            mPrint_Count = pPrint_Count;       //매수 
            mOutChoice = pOutChoice;           //출력방향[프린터/파일]
        }
        #endregion;

        #region ----- Export File Name Methods ----

        private string SetExportFileName(string pExportFileName)
        {
            string vExportFileName = string.Empty;

            try
            {
                vExportFileName = pExportFileName;
                vExportFileName = vExportFileName.Replace("/", "_");
                vExportFileName = vExportFileName.Replace("\\", "_");
                vExportFileName = vExportFileName.Replace("*", "_");
                vExportFileName = vExportFileName.Replace("<", "_");
                vExportFileName = vExportFileName.Replace(">", "_");
                vExportFileName = vExportFileName.Replace("|", "_");
                vExportFileName = vExportFileName.Replace("?", "_");
                vExportFileName = vExportFileName.Replace(":", "_");
                vExportFileName = vExportFileName.Replace(" ", "_");
            }
            catch
            {
            }

            return vExportFileName;
        }


        #endregion;

        #region ----- Private Methods -----

        private void HRMF0705_PRINT_Load(object sender, EventArgs e)
        {
            iedCORP_ID.EditValue = mCorp_ID;               //업체ID
            iedPRINT_NUM.EditValue = mPrint_Num;           //발급번호
            iedPRINT_DATE.EditValue = mPrint_Date;         //발급일자
            iedPRINT_YEAR.EditValue = mPrint_Year;         //징수년도            

            iedCERT_TYPE_NAME.EditValue = mCert_Type_Name; //증명서구분
            iedCERT_TYPE_ID.EditValue = mCert_Type_ID;     //증명서 구분ID
            iedCERT_TYPE_CODE.EditValue = mCert_Type_Code; //증명서 코드            
            iedPRINT_COUNT.EditValue = 1;                   //인쇄매수.

            this.Text = string.Format("{0} - {1}", this.Text, mOutChoice);

            if (mPerson_ID == null)
            {
                // 성명 란에 '전체'로 되어 있을 경우 증명서 발급 폼에 공백으로 처리
                iedNAME.EditValue = "";
            }
            else
            {
                iedNAME.EditValue = mName;                 //성명
            }

            iedPERSON_ID.EditValue = mPerson_ID;           //사원ID                        

            if (mJoin_Date.Year == 1)
            {
                iedJOIN_DATE.EditValue = DBNull.Value;     //입사일자
            }
            else
            {
                iedJOIN_DATE.EditValue = mJoin_Date;
            }

            if (mRetire_Date.Year == 1)
            {
                iedRETIRE_DATE.EditValue = DBNull.Value;   //퇴사일자
            }
            else
            {
                iedRETIRE_DATE.EditValue = mRetire_Date;
            }

            iedDESCRIPTION.EditValue = mDescription;       //용도
            iedSEND_ORG.EditValue = mSend_Org;             //제출처
            isPrintBt.CheckedState = ISUtil.Enum.CheckedState.Checked;
            // mOutChoice = "PRINT";

            //==============================================================
            //  갑종근로소득원천징수영수증
            //==============================================================
            PAY_YYYYMM_FR_0.EditValue = string.Format("{0}-01", iDate.ISYear(DateTime.Today));
            PAY_YYYYMM_TO_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            START_DATE_0.EditValue = iDate.ISMonth_1st(iDate.ISGetDate(PAY_YYYYMM_FR_0.EditValue));
            END_DATE_0.EditValue = iDate.ISMonth_Last(iDate.ISGetDate(PAY_YYYYMM_TO_0.EditValue));

            if (iString.ISNull(iedCERT_TYPE_CODE.EditValue) == "23")
            {
                PAY_YYYYMM_FR_0.Show();
                PAY_YYYYMM_TO_0.Show();
                iedPRINT_YEAR.Hide();
            }
            else
            {
                PAY_YYYYMM_FR_0.Hide();
                PAY_YYYYMM_TO_0.Hide();
                iedPRINT_YEAR.Show();
            }
            //==============================================================

            iedPRINT_DATE.Focus();
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

        #region ----- XL Export Methods ----

        private void ExportXL(ISDataAdapter pAdapter)
        {
            int vCountRow = pAdapter.OraSelectData.Rows.Count;
            if (vCountRow < 1)
            {
                return;
            }

            string vsMessage = string.Empty;
            string vsSheetName = "Slip_Line";

            saveFileDialog1.Title = "Excel_Save";
            saveFileDialog1.FileName = "XL_00";
            saveFileDialog1.DefaultExt = "xlsx";
            System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            saveFileDialog1.InitialDirectory = vSaveFolder.FullName;
            saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string vsSaveExcelFileName = saveFileDialog1.FileName;
                XL.XLPrint xlExport = new XL.XLPrint();
                bool vXLSaveOK = xlExport.XLExport(pAdapter.OraSelectData, vsSaveExcelFileName, vsSheetName);
                if (vXLSaveOK == true)
                {
                    vsMessage = string.Format("Save OK [{0}]", vsSaveExcelFileName);
                    MessageBox.Show(vsMessage);
                }
                else
                {
                    vsMessage = string.Format("Save Err [{0}]", vsSaveExcelFileName);
                    MessageBox.Show(vsMessage);
                }
                xlExport.XLClose();
            }
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

        #region ----- XL Print 2 Methods 20 -----
        //HRMF0705_002.xls
        private void XLPrinting_23(string pOutChoice)
        {
            System.DateTime vStartTime = DateTime.Now;
            string vMessageText = string.Empty;

            int vCountRow = gridPRINT_INCOME_TAX.RowCount; //gridWITHHOLDING_TAX 그리드의 총 행수
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            iedPRINT_DATE.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();



            XLPrinting_2 xlPrinting = new XLPrinting_2(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                vMessageText = string.Format(" XL Opening...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                string vDate = iedPRINT_DATE.DateTimeValue.ToString("yyyy 년  MM 월  dd 일", null);
                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0705_002.xlsx";
                //-------------------------------------------------------------------------------------

                string vSend_ORG = string.Format("{0}", iedSEND_ORG.EditValue);
                string vPrint_COUNT = string.Format("{0}", iedPRINT_COUNT.EditValue);

                object vPrintDate = iedPRINT_DATE.DateTimeValue.ToString("yyyy 년  MM 월  dd 일", null);
                //---------------------------------------------------------------------
                // 출력 용도 구분
                //---------------------------------------------------------------------
                object vPrintType = null;

                bool isOpen = xlPrinting.XLFileOpen();
                if (isOpen == true)
                {
                    vPageNumber = xlPrinting.WriteMain(gridPRINT_INCOME_TAX, vPrintDate, vPrintType, vSend_ORG, vPrint_COUNT, vDate);
                }

                for (int nCnt = 1; nCnt <= Convert.ToInt32(iedPRINT_COUNT.EditValue); nCnt++)
                {
                    string vSaveFileName = string.Empty;

                    if (pOutChoice == "PDF")
                    {
                        System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                        //저장 파일 이름 : 

                        string vName = gridPRINT_INCOME_TAX.GetCellValue("NAME").ToString();

                        vSaveFileName = string.Format("{0}_{1}", "갑종근로소득세원천징수증명서", vName);
                        vSaveFileName = SetExportFileName(vSaveFileName);

                        System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);

                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }

                    }

                    if (mOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (mOutChoice == "PDF")
                    {
                        xlPrinting.DeleteSheet();
                        xlPrinting.PDF(vSaveFileName);
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


        #region ----- XL Print 3 Methods 21 -----
        //HRMF0705_003.xls
        private void XLPrinting_21(string pOutChoice)
        {
            System.DateTime vStartTime = DateTime.Now;
            string vMessageText = string.Empty;

            int vCountRow = gridIN_EARNER_DED_TAX.RowCount; //gridIN_EARNER_DED_TAX 그리드의 총 행수
            if (vCountRow < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            iedPRINT_DATE.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;

            vMessageText = string.Format(" Printing Starting...");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting_3 xlPrinting = new XLPrinting_3(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                vMessageText = string.Format(" XL Opening...");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();

                //-------------------------------------------------------------------------------------
                xlPrinting.OpenFileNameExcel = "HRMF0705_003.xlsx";
                //-------------------------------------------------------------------------------------

                vPageNumber = xlPrinting.WriteMain(gridIN_EARNER_DED_TAX);

                for (int nCnt = 1; nCnt <= Convert.ToInt32(iedPRINT_COUNT.EditValue); nCnt++)
                {
                    string vSaveFileName = string.Empty;

                    if (pOutChoice == "PDF")
                    {
                        System.IO.DirectoryInfo vSaveFolder = new System.IO.DirectoryInfo(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                        //저장 파일 이름 : 

                        string vName = gridIN_EARNER_DED_TAX.GetCellValue("NAME").ToString();

                        vSaveFileName = string.Format("{0}_{1}", "근로소득원천징수부", vName);
                        vSaveFileName = SetExportFileName(vSaveFileName);

                        System.IO.FileInfo vFileName = new System.IO.FileInfo(vSaveFileName);

                        if (vFileName.Exists)
                        {
                            vFileName.Delete();
                        }

                    }

                    if (mOutChoice == "PRINT")
                    {
                        xlPrinting.Printing(1, vPageNumber);
                    }
                    else if (mOutChoice == "PDF")
                    {
                        xlPrinting.DeleteSheet();
                        xlPrinting.PDF(vSaveFileName);
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


        #region ----- Button Event ----

        // 발급 버튼 선택
        private void btnPRINT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iedCERT_TYPE_ID.EditValue == null)
            {// 증명서 구분
                MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10033"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedCERT_TYPE_NAME.Focus();
                return;
            }

            if (string.IsNullOrEmpty(iedDESCRIPTION.EditValue.ToString()))
            {// 용도 입력
                MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10034"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                iedCERT_TYPE_NAME.Focus();
                return;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            #region ----- 근로소득원천징수부 -----

            if (Convert.ToInt32(iedCERT_TYPE_CODE.EditValue) == 21) //근로소득원천징수부
            {
                if (iedPERSON_ID.EditValue == null)
                {// 성명 입력
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=1인 출력만 가능하므로 '성명'"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    iedNAME.Focus();
                    return;
                }

                if (CB_PRINT_SAVING_YN.CheckBoxString == "Y" || CB_PRINT_HOUSE_YN.CheckBoxString == "Y")
                {
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    CB_PRINT_SAVING_YN.CheckBoxValue = "N";
                    CB_PRINT_HOUSE_YN.CheckBoxValue = "N";
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("HRM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 인쇄 결과 저장
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_CORP_ID", iedCORP_ID.EditValue);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                idcCERTIFICATE_PRINT_INSERT.ExecuteNonQuery();
                iedPRINT_NUM.EditValue = idcCERTIFICATE_PRINT_INSERT.GetCommandParamValue("P_PRINT_NUM");
                idaIN_EARNER_DED_TAX.Fill();

                XLPrinting_21(mOutChoice); // 출력 함수 호출

                //// 인쇄발급 루틴 추가
                if (iString.ISNull(iedPRINT_NUM.EditValue) == string.Empty)
                {// 인쇄번호 없음. 인쇄 실패.
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10172"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10035"), "", MessageBoxButtons.OK, MessageBoxIcon.None);
                // 인쇄 완료 메시지 출력

                // EditBox 초기화
                iedPRINT_NUM.EditValue = null;              //발급번호
                iedNAME.EditValue = null;                   //성명
                iedPERSON_NUM.EditValue = null;             //사번
                iedPERSON_ID.EditValue = null;              //사원ID
                iedJOIN_DATE.EditValue = null;              //입사일자 
                iedRETIRE_DATE.EditValue = null;            //퇴사일자
                iedDESCRIPTION.EditValue = null;            //용도
                iedSEND_ORG.EditValue = null;               //제출처
                iedPRINT_COUNT.EditValue = 1;               //매수
                //iedCERT_TYPE_NAME.EditValue = null;       //증명서 이름
                //iedCERT_TYPE_ID.EditValue = null;         //증명서 ID
                //iedCERT_TYPE_CODE.EditValue = null;       //증명서 코드
                // 인쇄용도 구분 초기화
                EARNER_YN.CheckBoxValue = "N";
                ADDRESSOR1_YN.CheckBoxValue = "N";
                ADDRESSOR2_YN.CheckBoxValue = "N";
            }

            #endregion;

            #region ----- 근로소득영수증 ----

            else if (Convert.ToInt32(iedCERT_TYPE_CODE.EditValue) == 22) // 근로소득영수증
            {
                //---------------------------------------------------------------------
                // 출력 용도 구분 체크
                //---------------------------------------------------------------------

                if (EARNER_YN.CheckBoxString != "Y" && ADDRESSOR1_YN.CheckBoxString != "Y" && ADDRESSOR2_YN.CheckBoxString != "Y")
                {
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    MessageBox.Show("출력 용도를 선택해주세요.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    EARNER_YN.Focus();
                    return;
                }

                //// 인쇄 결과 저장
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_CORP_ID", iedCORP_ID.EditValue);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                idcCERTIFICATE_PRINT_INSERT.ExecuteNonQuery();
                iedPRINT_NUM.EditValue = idcCERTIFICATE_PRINT_INSERT.GetCommandParamValue("P_PRINT_NUM");

                DialogResult vdlgResult;

                Form vHRMF0705_PRINT_22 = new HRMF0705_PRINT_22(isAppInterfaceAdv1.AppInterface
                                                              , iedCORP_ID.EditValue                       //업체ID
                                                              , iedPERSON_ID.EditValue                     //사원ID
                                                              , iedPRINT_YEAR.EditValue                    //징수년도   
                                                              , JOB_CATEGORY_ID.EditValue
                                                              , FLOOR_ID.EditValue
                                                              , CB_EMPLOYE_3_YN.CheckBoxValue
                                                              , CB_PRINT_SAVING_YN.CheckBoxValue
                                                              , CB_PRINT_HOUSE_YN.CheckBoxValue
                                                              , iedPRINT_DATE.EditValue
                                                              , EARNER_YN.CheckBoxString
                                                              , ADDRESSOR1_YN.CheckBoxString
                                                              , ADDRESSOR2_YN.CheckBoxString
                                                              , mOutChoice
                                                              );
                vdlgResult = vHRMF0705_PRINT_22.ShowDialog();

                if (vdlgResult == DialogResult.OK)
                { 
                }
                vHRMF0705_PRINT_22.Dispose();


                //// 인쇄발급 루틴 추가
                if (iString.ISNull(iedPRINT_NUM.EditValue) == string.Empty)
                {// 인쇄번호 없음. 인쇄 실패.
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10172"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10035"), "", MessageBoxButtons.OK, MessageBoxIcon.None);
                // 인쇄 완료 메시지 출력

                //// EditBox 초기화
                iedPRINT_NUM.EditValue = null;              //발급번호
                iedNAME.EditValue = null;                   //성명
                iedPERSON_NUM.EditValue = null;             //사번
                iedPERSON_ID.EditValue = null;              //사원ID
                iedJOIN_DATE.EditValue = null;              //입사일자 
                iedRETIRE_DATE.EditValue = null;            //퇴사일자
                iedDESCRIPTION.EditValue = null;            //용도
                iedSEND_ORG.EditValue = null;               //제출처
                iedPRINT_COUNT.EditValue = 1;               //매수
                //////iedCERT_TYPE_NAME.EditValue = null;         //증명서 이름
                //////iedCERT_TYPE_ID.EditValue = null;           //증명서 ID
                //////iedCERT_TYPE_CODE.EditValue = null;         //증명서 코드

                // 인쇄용도 구분 초기화
                EARNER_YN.CheckBoxValue = "N";
                ADDRESSOR1_YN.CheckBoxValue = "N";
                ADDRESSOR2_YN.CheckBoxValue = "N";
            }

            #endregion;

            #region ----- 갑종근로소득에대한소득세원천징수증명서 -----

            else if (Convert.ToInt32(iedCERT_TYPE_CODE.EditValue) == 23) //갑종근로소득에대한소득세원천징수증명서            
            {
                if (iedPERSON_ID.EditValue == null)
                {// 성명 입력
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=1인 출력만 가능하므로 '성명'"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    iedNAME.Focus();
                    return;
                }

                if (CB_PRINT_SAVING_YN.CheckBoxString == "Y" || CB_PRINT_HOUSE_YN.CheckBoxString == "Y")
                {
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    CB_PRINT_SAVING_YN.CheckBoxValue = "N";
                    CB_PRINT_HOUSE_YN.CheckBoxValue = "N";
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("HRM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //// 인쇄 결과 저장
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_CORP_ID", iedCORP_ID.EditValue);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                idcCERTIFICATE_PRINT_INSERT.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                idcCERTIFICATE_PRINT_INSERT.ExecuteNonQuery();
                iedPRINT_NUM.EditValue = idcCERTIFICATE_PRINT_INSERT.GetCommandParamValue("P_PRINT_NUM");

                ida_PRINT_INCOME_TAX.Fill();

                XLPrinting_23(mOutChoice); // 출력 함수 호출

                //// 인쇄발급 루틴 추가
                if (iString.ISNull(iedPRINT_NUM.EditValue) == string.Empty)
                {// 인쇄번호 없음. 인쇄 실패.
                    Application.UseWaitCursor = false;
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                    Application.DoEvents();

                    MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10172"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                MessageBox.Show(isMessageAdapter1.ReturnText("FCM_10035"), "", MessageBoxButtons.OK, MessageBoxIcon.None);
                // 인쇄 완료 메시지 출력

                // EditBox 초기화
                iedPRINT_NUM.EditValue = null;              //발급번호
                iedNAME.EditValue = null;                   //성명
                iedPERSON_NUM.EditValue = null;             //사번
                iedPERSON_ID.EditValue = null;              //사원ID
                iedJOIN_DATE.EditValue = null;              //입사일자 
                iedRETIRE_DATE.EditValue = null;            //퇴사일자
                iedDESCRIPTION.EditValue = null;            //용도
                iedSEND_ORG.EditValue = null;               //제출처
                iedPRINT_COUNT.EditValue = 1;               //매수
                //iedCERT_TYPE_NAME.EditValue = null;         //증명서 이름
                //iedCERT_TYPE_ID.EditValue = null;           //증명서 ID
                //iedCERT_TYPE_CODE.EditValue = null;         //증명서 코드
                // 인쇄용도 구분 초기화
                EARNER_YN.CheckBoxValue = "N";
                ADDRESSOR1_YN.CheckBoxValue = "N";
                ADDRESSOR2_YN.CheckBoxValue = "N";
            }

            #endregion;

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        // 취소 버튼 선택
        private void btnCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            this.Close();
        }

        #endregion;

        #region ----- RowData Event -----

        private void ilaCERT_TYPE_SelectedRowData(object pSender)
        {
            if (iString.ISNull(iedCERT_TYPE_CODE.EditValue) == "23")
            {
                PAY_YYYYMM_FR_0.EditValue = string.Format("{0}-01", iDate.ISYear(DateTime.Today));
                PAY_YYYYMM_TO_0.EditValue = iDate.ISYearMonth(DateTime.Today);
                START_DATE_0.EditValue = iDate.ISMonth_1st(iDate.ISGetDate(PAY_YYYYMM_FR_0.EditValue));
                END_DATE_0.EditValue = iDate.ISMonth_Last(iDate.ISGetDate(PAY_YYYYMM_TO_0.EditValue));

                PAY_YYYYMM_FR_0.Show(); //징수년월(시작)
                PAY_YYYYMM_TO_0.Show(); //징수년월(종료)
                iedPRINT_YEAR.Hide();   //징수년도
            }
            else
            {
                PAY_YYYYMM_FR_0.Hide(); //징수년월(시작)
                PAY_YYYYMM_TO_0.Hide(); //징수년월(종료)
                iedPRINT_YEAR.Show();   //징수년도
            }
        }

        #endregion;

        #region ----- Lookup Event -----

        private void ilaYEAR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYEAR.SetLookupParamValue("W_START_YEAR", "2001");
            ildYEAR.SetLookupParamValue("W_END_YEAR", iDate.ISYear(DateTime.Today));
        }

        private void ilaCERT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "CERT_TYPE");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", "HC.VALUE1 = 20 ");
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_JOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (iedCERT_TYPE_CODE.EditValue.ToString() == "23".ToString())
            {
                ildPERSON.SetLookupParamValue("W_START_DATE", START_DATE_0.EditValue);
                ildPERSON.SetLookupParamValue("W_END_DATE", END_DATE_0.EditValue);
            }
            else
            {
                ildPERSON.SetLookupParamValue("W_START_DATE", iDate.ISGetDate(String.Format("{0}-01", iedPRINT_YEAR.EditValue)));
                ildPERSON.SetLookupParamValue("W_END_DATE", iDate.ISGetDate(String.Format("{0}-12-31", iedPRINT_YEAR.EditValue)));
            }
            ildPERSON.SetLookupParamValue("W_CORP_ID", iedCORP_ID.EditValue);
        }

        private void ilaYYYYMM_FR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            string sEndDate = string.Format("{0}{1}", iDate.ISYear(DateTime.Today), "-12");
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2000-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", sEndDate);
        }

        private void ilaYYYYMM_TO_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            string sEndDate = string.Format("{0}{1}", iDate.ISYear(DateTime.Today), "-12");
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2000-01");
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", sEndDate);
        }

        private void isRadioButtonAdv_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;

            if (vRadio.Checked == true)
            {
                mOutChoice = vRadio.RadioCheckedString;
            }

        }
        #endregion;


    }
}