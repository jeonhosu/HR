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

namespace HRMF0201
{
    public partial class HRMF0201 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iString = new ISFunction.ISConvert();

        private string mMessageError = string.Empty;
        private string mCARD_VALUE = string.Empty;
        bool mSUB_SHOW_FLAG = false;

        #endregion;

        #region ----- UpLoad / DownLoad Variables -----

        private InfoSummit.Win.ControlAdv.ISFileTransferAdv mFileTransferAdv;
        private ItemImageInfomationFTP mImageFTP;
        private PersonDocFTP mDocFTP;

        private string mFTP_Source_Directory = string.Empty;            // ftp 소스 디렉토리.
        private string mClient_Base_Path = System.Windows.Forms.Application.StartupPath;    // 현재 디렉토리.

        private string mClient_Directory = string.Empty;            // 실제 디렉토리 
        private string mClient_ImageDirectory = string.Empty;       // 클라이언트 이미지 디렉토리.
        private string mFileExtension = ".JPG";                     // 확장자명.

        private string mClient_DocDirectory = string.Empty;         // 클라이언트 문서 디렉토리.
        private bool mIsGetInformationFTP = false;                  // FTP 정보 상태.
        private bool mIsGetPersonDocFTP = false;                    // 사원 정보 FTP 정보 상태.
        private bool mIsFormLoad = false;                           // NEWMOVE 이벤트 제어.

        private string mPerson_ImageLocation;                       //이미지 로케이션.

        bool mIsClickInquiryDetail = false;
        int mInquiryDetailPreX, mInquiryDetailPreY; //마우스 이동 제어.

        #endregion;

        #region ----- initialize -----

        public HRMF0201(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.DoubleBuffered = true;
            this.Visible = false;
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            mIsFormLoad = false;
        }
        #endregion
                
        #region ----- DATA FIND ------
        
        private void SEARCH_DB()
        {// 데이터 조회.

            string vPERSON_NUM = iString.ISNull(igrPERSON_INFO.GetCellValue("PERSON_NUM"));

            IDA_PERSON.Fill();

            int vIDX_PERSON_NUM = igrPERSON_INFO.GetColumnToIndex("PERSON_NUM");
            if (vPERSON_NUM != string.Empty)
            {
                for (int i = 0; i < igrPERSON_INFO.RowCount; i++)
                {
                    if (iString.ISNull(igrPERSON_INFO.GetCellValue(i, vIDX_PERSON_NUM)) == vPERSON_NUM)
                    {
                        igrPERSON_INFO.CurrentCellMoveTo(i, vIDX_PERSON_NUM);
                        igrPERSON_INFO.Focus();
                        return;
                    }
                }
            }
            igrPERSON_INFO.Focus();
        }

        private void isSearch_Sub_DB()
        {// 서브 tab 조회.
            if (PERSON_ID.EditValue == null)
            {
                return;
            }
            idaBODY.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaBODY.Fill();

            idaARMY.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaARMY.Fill();

            idaFAMILY.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            idaFAMILY.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            idaFAMILY.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaFAMILY.Fill();

            idaHISTORY.SetSelectParamValue("W_HISTORY_HEADER_ID", DBNull.Value);
            idaHISTORY.SetSelectParamValue("W_DEPT_ID", DBNull.Value);
            idaHISTORY.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaHISTORY.Fill();

            idaCAREER.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaCAREER.Fill();

            idaSCHOLARSHIP.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaSCHOLARSHIP.Fill();

            idaEDUCATION.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaEDUCATION.Fill();

            idaRESULT.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaRESULT.Fill();

            idaLICENSE.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaLICENSE.Fill();

            idaFOREIGN_LANGUAGE.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaFOREIGN_LANGUAGE.Fill();

            idaREWARD_PUNISHMENT.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaREWARD_PUNISHMENT.Fill();

            idaREFERENCE.SetSelectParamValue("W_PERSON_ID", PERSON_ID.EditValue);
            idaREFERENCE.Fill();

        }
        #endregion

        #region ----- Data validate -----
        private bool isPerson_ID_Validate()
        {// 사원번호 존재 여부 체크.
            bool ibReturn_Value = false;
            if (PERSON_ID.EditValue == null)
            {
                ibReturn_Value = false;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK,MessageBoxIcon.Warning);  // 사원정보는 필수.
            }
            else
            {
                ibReturn_Value = true;
            }
            return ibReturn_Value;
        }

        #endregion

        #region ----- 생년월일 생성 -----

        private DateTime BIRTHDAY(object pREPRE_NUM)
        {
            DateTime mBIRTHDAY;

            string mSex_Type = pREPRE_NUM.ToString().Replace("-", "").Substring(6, 1);
            if (mSex_Type == "1".ToString() || mSex_Type == "2".ToString() || mSex_Type == "5".ToString() || mSex_Type == "6".ToString())
            {
                mBIRTHDAY = DateTime.Parse("19" + pREPRE_NUM.ToString().Substring(0, 2)
                                                    + "-".ToString()
                                                    + pREPRE_NUM.ToString().Substring(2, 2)
                                                    + "-".ToString()
                                                    + pREPRE_NUM.ToString().Substring(4, 2));
            }
            else
            {
                mBIRTHDAY = DateTime.Parse("20" + pREPRE_NUM.ToString().Substring(0, 2)
                                                    + "-".ToString()
                                                    + pREPRE_NUM.ToString().Substring(2, 2)
                                                    + "-".ToString()
                                                    + pREPRE_NUM.ToString().Substring(4, 2));
            }
            return mBIRTHDAY;
        }

        #endregion

        #region ----- 주민번호 체크 ------

        private void iedREPRE_NUM_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            if (string.IsNullOrEmpty(iedREPRE_NUM.EditValue.ToString()))
            {
                return;
            }

            // 전호수 주석 : '-' 입력 체크 안함. 단, DB에서 자릿수 검증후 '-' 자동 입력 처리.
            //if (iedREPRE_NUM.EditValue.ToString().IndexOf("-") == -1)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //}

            string isReturnValue = null;
            idcREPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", iedREPRE_NUM.EditValue);
            idcREPRE_NUM_CHECK.ExecuteNonQuery();
            isReturnValue = idcREPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE").ToString();
            iedSEX_TYPE.EditValue = idcREPRE_NUM_CHECK.GetCommandParamValue("O_SEX_TYPE");
            if (isReturnValue == "N".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10026"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
            if (string.IsNullOrEmpty(iedSEX_TYPE.EditValue.ToString()))
            {
                iedSEX_NAME.EditValue = null;
                iedSEX_TYPE.EditValue = null;
                return;
            }
            idcSEX_TYPE.OraProcedure = "CODE_NAME";
            idcSEX_TYPE.SetCommandParamValue("W_GROUP_CODE", "SEX_TYPE");
            idcSEX_TYPE.SetCommandParamValue("W_CODE", iedSEX_TYPE.EditValue);
            idcSEX_TYPE.ExecuteNonQuery();
            iedSEX_NAME.EditValue = idcSEX_TYPE.GetCommandParamValue("O_RETURN_VALUE").ToString();

            if (iString.ISNull(iedBIRTHDAY.EditValue) == string.Empty)
            {// 생년월일이 기존에 없을 경우 자동 설정.
                iedBIRTHDAY.EditValue = BIRTHDAY(iedREPRE_NUM.EditValue);

                // 음양구분.
                idcCOMMON_W.SetCommandParamValue("W_GROUP_CODE", "BIRTHDAY_TYPE");
                idcCOMMON_W.SetCommandParamValue("W_WHERE", " 1 = 1 ");
                idcCOMMON_W.ExecuteNonQuery();
                iedBIRTHDAY_TYPE_NAME.EditValue = idcCOMMON_W.GetCommandParamValue("O_CODE_NAME");
                iedBIRTHDAY_TYPE.EditValue = idcCOMMON_W.GetCommandParamValue("O_CODE");
            }
        }

        private string FAMILY_REPRE_NUM_CHECK(object pREPRE_NUM)
        {
            string isReturnValue = "N".ToString();
            if (iString.ISNull(pREPRE_NUM) == string.Empty)
            {
                return isReturnValue;
            }

            // 전호수 주석 : '-' 입력 체크 안함. 단, DB에서 자릿수 검증후 '-' 자동 입력 처리.
            //if (iedREPRE_NUM.EditValue.ToString().IndexOf("-") == -1)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return isReturnValue;
            //}
    
            idcREPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", pREPRE_NUM);
            idcREPRE_NUM_CHECK.ExecuteNonQuery();
            isReturnValue = idcREPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE").ToString();
            return isReturnValue;            
        }

        private void iedR_GUAR_REPRE_NUM1_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            string isReturnValue = null;
            if (iString.ISNull(iedR_GUAR_REPRE_NUM1.EditValue) == string.Empty)
            {
                return;
            }
            idcREPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", iedR_GUAR_REPRE_NUM1.EditValue);
            idcREPRE_NUM_CHECK.ExecuteNonQuery();
            isReturnValue = idcREPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE").ToString();
            if (isReturnValue == "N".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10026"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
        }

        private void iedR_GUAR_REPRE_NUM2_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            string isReturnValue = null;
            if (iString.ISNull(iedR_GUAR_REPRE_NUM2.EditValue) == string.Empty)
            {
                return;
            }
            idcREPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", iedR_GUAR_REPRE_NUM2.EditValue);
            idcREPRE_NUM_CHECK.ExecuteNonQuery();
            isReturnValue = idcREPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE").ToString();            
            if (isReturnValue == "N".ToString())
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10026"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
        }
        
        #endregion

        #region ----- 주소 조회 -----

        private void Show_Address_Legal()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, iedLEGAL_ZIP_CODE.EditValue, iedLEGAL_ADDR1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                iedLEGAL_ZIP_CODE.EditValue = vEAPF0299.Get_Zip_Code;
                iedLEGAL_ADDR1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_PRSN()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
            
            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, iedPRSN_ZIP_CODE.EditValue, iedPRSN_ADDR1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                iedPRSN_ZIP_CODE.EditValue = vEAPF0299.Get_Zip_Code;
                iedPRSN_ADDR1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_Live()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
            
            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, iedLIVE_ZIP_CODE.EditValue, iedLIVE_ADDR1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                iedLIVE_ZIP_CODE.EditValue = vEAPF0299.Get_Zip_Code;
                iedLIVE_ADDR1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_Career(int pIDX_Row, int pIDX_ZIP_CODE, int pIDX_ADDRESS)
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            igrCAREER.LastConfirmChanges();
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent
                                                    , isAppInterfaceAdv1.AppInterface
                                                    , igrCAREER.GetCellValue(pIDX_Row, pIDX_ZIP_CODE)
                                                    , igrCAREER.GetCellValue(pIDX_Row, pIDX_ADDRESS));
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                igrCAREER.SetCellValue(pIDX_Row, pIDX_ZIP_CODE, vEAPF0299.Get_Zip_Code);
                igrCAREER.SetCellValue(pIDX_Row, pIDX_ADDRESS, vEAPF0299.Get_Address);
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_GUAR1()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
            
            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, iedR_GUAR_ZIP_CODE1.EditValue, iedR_GUAR_ADDR1_1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                iedR_GUAR_ZIP_CODE1.EditValue = vEAPF0299.Get_Zip_Code;
                iedR_GUAR_ADDR1_1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_GUAR2()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
            
            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, iedR_GUAR_ZIP_CODE2.EditValue, iedR_GUAR_ADDR2_1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                iedR_GUAR_ZIP_CODE2.EditValue = vEAPF0299.Get_Zip_Code;
                iedR_GUAR_ADDR2_1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        #endregion

        #region  ------ Property / Method -----

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
            ildCORP.SetLookupParamValue("W_DEPT_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            W_CORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            W_CORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            W_CORP_NAME.BringToFront();

            //재직구분
            idcDEFAULT_VALUE_GROUP.SetCommandParamValue("W_GROUP_CODE", "EMPLOYE_TYPE");
            idcDEFAULT_VALUE_GROUP.ExecuteNonQuery();
            W_EMPLOYE_TYPE.EditValue = idcDEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE");
            W_EMPLOYE_TYPE_NAME.EditValue = idcDEFAULT_VALUE_GROUP.GetCommandParamValue("O_CODE_NAME");

            //카드번호 연동하기 위한 기본값//
            IDC_GET_IC_CARD_VALUE_P.ExecuteNonQuery();
            mCARD_VALUE = iString.ISNull(IDC_GET_IC_CARD_VALUE_P.GetCommandParamValue("O_CARD_VALUE"));
        }

        private void isSetCommonLookUpParameter(string P_GROUP_CODE, string P_CODE_NAME, String P_USABLE_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", P_GROUP_CODE);
            ildCOMMON.SetLookupParamValue("W_CODE_NAME", P_CODE_NAME);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", P_USABLE_YN);
        }        

        private void Init_Person_Insert()
        {// 인사정보 insert.
            //iedORI_JOIN_DATE.EditValue = iDate.ISGetDate();
            //iedJOIN_DATE.EditValue = iDate.ISGetDate();

            if (V_PERSON_COPY.CheckedState == ISUtil.Enum.CheckedState.Checked)
            {
                int mPreRowPosition = IDA_PERSON.CurrentRowPosition() - 1;
                if (mPreRowPosition > -1)
                {
                    igrPERSON_INFO.SetCellValue("CORP_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["CORP_ID"]);
                    igrPERSON_INFO.SetCellValue("CORP_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["CORP_NAME"]);
                    igrPERSON_INFO.SetCellValue("OPERATING_UNIT_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["OPERATING_UNIT_ID"]);
                    igrPERSON_INFO.SetCellValue("OPERATING_UNIT_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["OPERATING_UNIT_NAME"]);
                    igrPERSON_INFO.SetCellValue("DEPT_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["DEPT_ID"]);
                    igrPERSON_INFO.SetCellValue("DEPT_CODE", IDA_PERSON.CurrentRows[mPreRowPosition]["DEPT_CODE"]);
                    igrPERSON_INFO.SetCellValue("DEPT_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["DEPT_NAME"]);
                    igrPERSON_INFO.SetCellValue("NATION_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["NATION_ID"]);
                    igrPERSON_INFO.SetCellValue("NATION_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["NATION_NAME"]);
                    igrPERSON_INFO.SetCellValue("WORK_AREA_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["WORK_AREA_ID"]);
                    igrPERSON_INFO.SetCellValue("WORK_AREA_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["WORK_AREA_NAME"]);
                    igrPERSON_INFO.SetCellValue("WORK_TYPE_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["WORK_TYPE_ID"]);
                    igrPERSON_INFO.SetCellValue("WORK_TYPE_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["WORK_TYPE_NAME"]);
                    igrPERSON_INFO.SetCellValue("JOB_CLASS_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["JOB_CLASS_ID"]);
                    igrPERSON_INFO.SetCellValue("JOB_CLASS_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["JOB_CLASS_NAME"]);
                    igrPERSON_INFO.SetCellValue("JOB_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["JOB_ID"]);
                    igrPERSON_INFO.SetCellValue("JOB_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["JOB_NAME"]);
                    igrPERSON_INFO.SetCellValue("POST_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["POST_ID"]);
                    igrPERSON_INFO.SetCellValue("POST_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["POST_NAME"]);
                    igrPERSON_INFO.SetCellValue("OCPT_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["OCPT_ID"]);
                    igrPERSON_INFO.SetCellValue("OCPT_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["OCPT_NAME"]);
                    igrPERSON_INFO.SetCellValue("ABIL_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["ABIL_ID"]);
                    igrPERSON_INFO.SetCellValue("ABIL_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["ABIL_NAME"]);
                    igrPERSON_INFO.SetCellValue("PAY_GRADE_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["PAY_GRADE_ID"]);
                    igrPERSON_INFO.SetCellValue("PAY_GRADE_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["PAY_GRADE_NAME"]);
                    igrPERSON_INFO.SetCellValue("BIRTHDAY_TYPE", IDA_PERSON.CurrentRows[mPreRowPosition]["BIRTHDAY_TYPE"]);
                    igrPERSON_INFO.SetCellValue("BIRTHDAY_TYPE_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["BIRTHDAY_TYPE_NAME"]);                    
                    igrPERSON_INFO.SetCellValue("CONTRACT_TYPE_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["CONTRACT_TYPE_ID"]);
                    igrPERSON_INFO.SetCellValue("CONTRACT_TYPE_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["CONTRACT_TYPE_NAME"]);
                    igrPERSON_INFO.SetCellValue("JOIN_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["JOIN_ID"]);
                    igrPERSON_INFO.SetCellValue("JOIN_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["JOIN_NAME"]);
                    igrPERSON_INFO.SetCellValue("JOIN_ROUTE_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["JOIN_ROUTE_ID"]);
                    igrPERSON_INFO.SetCellValue("JOIN_ROUTE_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["JOIN_ROUTE_NAME"]);
                    igrPERSON_INFO.SetCellValue("ORI_JOIN_DATE", IDA_PERSON.CurrentRows[mPreRowPosition]["ORI_JOIN_DATE"]);
                    igrPERSON_INFO.SetCellValue("JOIN_DATE", IDA_PERSON.CurrentRows[mPreRowPosition]["JOIN_DATE"]);
                    igrPERSON_INFO.SetCellValue("PAY_DATE", IDA_PERSON.CurrentRows[mPreRowPosition]["PAY_DATE"]);
                    igrPERSON_INFO.SetCellValue("DIR_INDIR_TYPE", IDA_PERSON.CurrentRows[mPreRowPosition]["DIR_INDIR_TYPE"]);
                    igrPERSON_INFO.SetCellValue("DIR_INDIR_TYPE_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["DIR_INDIR_TYPE_NAME"]);                     
                    igrPERSON_INFO.SetCellValue("END_SCH_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["END_SCH_ID"]);
                    igrPERSON_INFO.SetCellValue("END_SCH_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["END_SCH_NAME"]);
                    igrPERSON_INFO.SetCellValue("JOB_CATEGORY_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["JOB_CATEGORY_ID"]);
                    igrPERSON_INFO.SetCellValue("JOB_CATEGORY_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["JOB_CATEGORY_NAME"]);
                    igrPERSON_INFO.SetCellValue("FLOOR_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["FLOOR_ID"]);
                    igrPERSON_INFO.SetCellValue("FLOOR_NAME", IDA_PERSON.CurrentRows[mPreRowPosition]["FLOOR_NAME"]);
                    igrPERSON_INFO.SetCellValue("COST_CENTER_ID", IDA_PERSON.CurrentRows[mPreRowPosition]["COST_CENTER_ID"]);
                    igrPERSON_INFO.SetCellValue("COST_CENTER", IDA_PERSON.CurrentRows[mPreRowPosition]["COST_CENTER"]);
                    igrPERSON_INFO.SetCellValue("CORP_TYPE", IDA_PERSON.CurrentRows[mPreRowPosition]["CORP_TYPE"]);
                    igrPERSON_INFO.SetCellValue("LABOR_UNION_YN", IDA_PERSON.CurrentRows[mPreRowPosition]["LABOR_UNION_YN"]);

                    igrPERSON_INFO.Invalidate();
                }
            }
            else
            {
                // LOOKUP DEFAULT VALUE SETTING - CORP
                idcDEFAULT_CORP.SetCommandParamValue("W_DEPT_CONTROL_YN", "Y");
                idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
                idcDEFAULT_CORP.ExecuteNonQuery();

                iedCORP_NAME.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                iedCORP_ID.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
                iedCORP_TYPE.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_TYPE");

                igrPERSON_INFO.SetCellValue("ORI_JOIN_DATE", DBNull.Value);
                igrPERSON_INFO.SetCellValue("JOIN_DATE", DBNull.Value);
                igrPERSON_INFO.SetCellValue("PAY_DATE", DBNull.Value);
                igrPERSON_INFO.SetCellValue("EXPIRE_DATE", DBNull.Value);
                igrPERSON_INFO.SetCellValue("RETIRE_DATE", DBNull.Value);
                igrPERSON_INFO.SetCellValue("MARRY_DATE", DBNull.Value); 
                igrPERSON_INFO.SetCellValue("LABOR_UNION_DATE", DBNull.Value); 
            }
            iedNAME.Focus();
        }

        private void Sub_Form_Visible(bool pShow_Flag, string pSub_Panel)
        {
            if (mSUB_SHOW_FLAG == true && pShow_Flag == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (pShow_Flag == true)
            {
                try
                {
                    if (pSub_Panel == "IC_CARD")
                    {
                        GB_IC_NUM.Left = 251;
                        GB_IC_NUM.Top = 277;

                        GB_IC_NUM.Width = 584;
                        GB_IC_NUM.Height = 197;

                        GB_IC_NUM.Border3DStyle = Border3DStyle.Bump;
                        GB_IC_NUM.BorderStyle = BorderStyle.Fixed3D;

                        GB_IC_NUM.BringToFront();
                        GB_IC_NUM.Visible = true;
                    }
                    else if (pSub_Panel == "PERSON_DOC")
                    {
                        GB_DOC_ATT.Left = 340;
                        GB_DOC_ATT.Top = 100;

                        GB_DOC_ATT.Width = 415;
                        GB_DOC_ATT.Height = 230;

                        GB_DOC_ATT.Border3DStyle = Border3DStyle.Bump;
                        GB_DOC_ATT.BorderStyle = BorderStyle.Fixed3D;

                        GB_DOC_ATT.Controls[0].MouseDown += GB_DOC_ATT_MouseDown;
                        GB_DOC_ATT.Controls[0].MouseMove += GB_DOC_ATT_MouseMove;
                        GB_DOC_ATT.Controls[0].MouseUp += GB_DOC_ATT_MouseUp;
                        GB_DOC_ATT.Controls[1].MouseDown += GB_DOC_ATT_MouseDown;
                        GB_DOC_ATT.Controls[1].MouseMove += GB_DOC_ATT_MouseMove;
                        GB_DOC_ATT.Controls[1].MouseUp += GB_DOC_ATT_MouseUp;
                        GB_DOC_ATT.BringToFront();
                        GB_DOC_ATT.Visible = true;
                    }
                    mSUB_SHOW_FLAG = true;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                GB_CONDITION.Enabled = false;
                SC_MAIN.Enabled = false;
            }
            else
            {
                try
                {
                    if (pSub_Panel == "IC_CARD")
                    {
                        GB_IC_NUM.Visible = false;
                    }
                    else if (pSub_Panel == "PERSON_DOC")
                    {
                        GB_DOC_ATT.Visible = false;
                    }
                    else
                    {
                        GB_IC_NUM.Visible = false;
                        GB_DOC_ATT.Visible = false;
                    }
                    mSUB_SHOW_FLAG = false;
                }
                catch
                {
                    mSUB_SHOW_FLAG = false;
                }
                GB_CONDITION.Enabled = true;
                SC_MAIN.Enabled = true;
            } 
        }

        private void Init_Sub_Panel(bool pShow_Flag, string pSub_Panel)
        {
            
            
        }

        #endregion

        #region --- Application_MainButtonClick ---

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
                    if (IDA_PERSON.IsFocused)
                    {// 기본정보
                        IDA_PERSON.AddOver();
                        Init_Person_Insert();
                    }
                    else if (idaBODY.IsFocused)
                    {// 신체사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaBODY.AddOver();
                        iedB_PERSON_ID.EditValue = PERSON_ID.EditValue;      //사원id copy
                    }
                    else if (idaARMY.IsFocused)
                    {// 병역사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaARMY.AddOver();
                        iedA_PERSON_ID.EditValue = PERSON_ID.EditValue;      //사원id copy
                    }
                    else if (idaFAMILY.IsFocused)
                    {// 가족사항
                        if (isPerson_ID_Validate() == false)
                        {   
                            return;
                        }
                        idaFAMILY.AddOver();
                        igrFAMILY.SetCellValue("PERSON_ID", PERSON_ID.EditValue);    //사원id copy
                        igrFAMILY.SetCellValue("YEAR_ADJUST_YN", "Y");                  //정산여부. 
                    }
                    else if (idaCAREER.IsFocused)
                    {// 경력사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaCAREER.AddOver();
                        igrCAREER.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaSCHOLARSHIP.IsFocused)
                    {// 학력사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaSCHOLARSHIP.AddOver();
                        igrSCHOLARSHIP.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaEDUCATION.IsFocused)
                    {// 교육사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaEDUCATION.AddOver();
                        igrEDUCATION.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaRESULT.IsFocused)
                    {// 평가사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaRESULT.AddOver();
                        igrRESULT.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaLICENSE.IsFocused)
                    {// 자격사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaLICENSE.AddOver();
                        igrLICENSE.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaFOREIGN_LANGUAGE.IsFocused)
                    {// 어학사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaFOREIGN_LANGUAGE.AddOver();
                        igrFOREIGN_LANGUAGE.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaREWARD_PUNISHMENT.IsFocused)
                    {// 상벌사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaREWARD_PUNISHMENT.AddOver();
                        igrREWARD_PUNISHMENT.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaREFERENCE.IsFocused)
                    {// 신원보증
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaREFERENCE.AddOver();
                        iedR_PERSON_ID.EditValue = PERSON_ID.EditValue;      //사원id copy
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_PERSON.IsFocused)
                    {// 기본정보
                        IDA_PERSON.AddUnder();
                        Init_Person_Insert();
                    }
                    else if (idaBODY.IsFocused)
                    {// 신체사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaBODY.AddUnder();
                        iedB_PERSON_ID.EditValue = PERSON_ID.EditValue;      //사원id copy
                    }
                    else if (idaARMY.IsFocused)
                    {// 병역사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaARMY.AddUnder();
                        iedA_PERSON_ID.EditValue = PERSON_ID.EditValue;      //사원id copy
                    }
                    else if (idaFAMILY.IsFocused)
                    {// 가족사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaFAMILY.AddUnder();
                        igrFAMILY.SetCellValue("PERSON_ID", PERSON_ID.EditValue);    //사원id copy
                        igrFAMILY.SetCellValue("YEAR_ADJUST_YN", "Y");                  //정산여부. 
                    }
                    else if (idaCAREER.IsFocused)
                    {// 경력사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaCAREER.AddUnder();
                        igrCAREER.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaSCHOLARSHIP.IsFocused)
                    {// 학력사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaSCHOLARSHIP.AddUnder();
                        igrSCHOLARSHIP.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaEDUCATION.IsFocused)
                    {// 교육사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaEDUCATION.AddUnder();
                        igrEDUCATION.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaRESULT.IsFocused)
                    {// 평가사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaRESULT.AddUnder();
                        igrRESULT.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaLICENSE.IsFocused)
                    {// 자격사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaLICENSE.AddUnder();
                        igrLICENSE.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaFOREIGN_LANGUAGE.IsFocused)
                    {// 어학사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaFOREIGN_LANGUAGE.AddUnder();
                        igrFOREIGN_LANGUAGE.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaREWARD_PUNISHMENT.IsFocused)
                    {// 상벌사항
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaREWARD_PUNISHMENT.AddUnder();
                        igrREWARD_PUNISHMENT.SetCellValue("PERSON_ID", PERSON_ID.EditValue);      //사원id copy
                    }
                    else if (idaREFERENCE.IsFocused)
                    {// 신원보증
                        if (isPerson_ID_Validate() == false)
                        {
                            return;
                        }
                        idaREFERENCE.AddUnder();
                        iedR_PERSON_ID.EditValue = PERSON_ID.EditValue;      //사원id copy
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    try
                    {
                        IDA_PERSON.Update();
                    }
                    catch
                    {

                    }
                    //if (idaPERSON.IsFocused)
                    //{// 기본정보
                    //    idaPERSON.Update();
                    //}
                    //else if (idaBODY.IsFocused)
                    //{// 신체사항
                    //    idaBODY.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaBODY.Update();
                    //}
                    //else if (idaARMY.IsFocused)
                    //{// 병역사항
                    //    idaARMY.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaARMY.Update();
                    //}
                    //else if (idaFAMILY.IsFocused)
                    //{// 가족사항
                    //    idaFAMILY.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaFAMILY.Update();
                    //}
                    //else if (idaCAREER.IsFocused)
                    //{// 경력사항
                    //    idaCAREER.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaCAREER.Update();
                    //}
                    //else if (idaSCHOLARSHIP.IsFocused)
                    //{// 학력사항
                    //    idaSCHOLARSHIP.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaSCHOLARSHIP.Update();
                    //}
                    //else if (idaEDUCATION.IsFocused)
                    //{// 교육사항
                    //    idaEDUCATION.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaEDUCATION.Update();
                    //}
                    //else if (idaRESULT.IsFocused)
                    //{// 평가사항
                    //    idaRESULT.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaRESULT.Update();
                    //}
                    //else if (idaLICENSE.IsFocused)
                    //{// 자격사항
                    //    idaLICENSE.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaLICENSE.Update();
                    //}
                    //else if (idaFOREIGN_LANGUAGE.IsFocused)
                    //{// 어학사항
                    //    idaFOREIGN_LANGUAGE.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaFOREIGN_LANGUAGE.Update();
                    //}
                    //else if (idaREWARD_PUNISHMENT.IsFocused)
                    //{// 상벌사항
                    //    idaREWARD_PUNISHMENT.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaREWARD_PUNISHMENT.Update();
                    //}
                    //else if (idaREFERENCE.IsFocused)
                    //{// 신원보증
                    //    idaREFERENCE.SetInsertParamValue("P_PERSON_ID", iedPERSON_ID.EditValue);
                    //    idaREFERENCE.Update();
                    //}
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_PERSON.IsFocused)
                    {// 기본정보
                        IDA_PERSON.Cancel();
                    }
                    else if (idaBODY.IsFocused)
                    {// 신체사항
                        idaBODY.Cancel();
                    }
                    else if (idaARMY.IsFocused)
                    {// 병역사항
                        idaARMY.Cancel();
                    }
                    else if (idaFAMILY.IsFocused)
                    {// 가족사항
                        idaFAMILY.Cancel();
                    }
                    else if (idaCAREER.IsFocused)
                    {// 경력사항
                        idaCAREER.Cancel();
                    }
                    else if (idaSCHOLARSHIP.IsFocused)
                    {// 학력사항
                        idaSCHOLARSHIP.Cancel();
                    }
                    else if (idaEDUCATION.IsFocused)
                    {// 교육사항
                        idaEDUCATION.Cancel();
                    }
                    else if (idaRESULT.IsFocused)
                    {// 평가사항
                        idaRESULT.Cancel();
                    }
                    else if (idaLICENSE.IsFocused)
                    {// 자격사항
                        idaLICENSE.Cancel();
                    }
                    else if (idaFOREIGN_LANGUAGE.IsFocused)
                    {// 어학사항
                        idaFOREIGN_LANGUAGE.Cancel();
                    }
                    else if (idaREWARD_PUNISHMENT.IsFocused)
                    {// 상벌사항
                        idaREWARD_PUNISHMENT.Cancel();
                    }
                    else if (idaREFERENCE.IsFocused)
                    {// 신원보증
                        idaREFERENCE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_PERSON.IsFocused)
                    {// 기본정보
                        IDA_PERSON.Delete();
                    }
                    else if (idaBODY.IsFocused)
                    {// 신체사항
                        idaBODY.Delete();
                    }
                    else if (idaARMY.IsFocused)
                    {// 병역사항
                        idaARMY.Delete();
                    }
                    else if (idaFAMILY.IsFocused)
                    {// 가족사항
                        idaFAMILY.Delete();
                    }
                    else if (idaCAREER.IsFocused)
                    {// 경력사항
                        idaCAREER.Delete();
                    }
                    else if (idaSCHOLARSHIP.IsFocused)
                    {// 학력사항
                        idaSCHOLARSHIP.Delete();
                    }
                    else if (idaEDUCATION.IsFocused)
                    {// 교육사항
                        idaEDUCATION.Delete();
                    }
                    else if (idaRESULT.IsFocused)
                    {// 평가사항
                        idaRESULT.Delete();
                    }
                    else if (idaLICENSE.IsFocused)
                    {// 자격사항
                        idaLICENSE.Delete();
                    }
                    else if (idaFOREIGN_LANGUAGE.IsFocused)
                    {// 어학사항
                        idaFOREIGN_LANGUAGE.Delete();
                    }
                    else if (idaREWARD_PUNISHMENT.IsFocused)
                    {// 상벌사항
                        idaREWARD_PUNISHMENT.Delete();
                    }
                    else if (idaREFERENCE.IsFocused)
                    {// 신원보증
                        idaREFERENCE.Delete();
                    }
                }
            }
        }

        #endregion

        #region ----- Form Event -----

        private void HRMF0201_Load(object sender, EventArgs e)
        {
            this.Visible = true;
            mIsFormLoad = true;

            Sub_Form_Visible(false, "");
        }

        private void HRMF0201_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();

            mIsGetInformationFTP = GetInfomationFTP();
            if (mIsGetInformationFTP == true)
            {
                MakeDirectory("PERSON_PIC");
                FTPInitializtion("PERSON_PIC");
            }

            mIsGetPersonDocFTP = GetPersonDocFTP();
            if(mIsGetPersonDocFTP)
            {
                MakeDirectory("PERSON_DOC"); 
            } 

            mIsFormLoad = false;

            IDA_PERSON.FillSchema();
            IDA_PERSON_IC_CARD.FillSchema();
        }
        
        private void HRMF0201_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (mIsGetInformationFTP == true)
            {
                System.IO.DirectoryInfo vClient_ImageDirectory = new System.IO.DirectoryInfo(mClient_ImageDirectory);
                if (vClient_ImageDirectory.Exists == true)
                {
                    try
                    {
                        vClient_ImageDirectory.Delete(true);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private void iedREPRE_NUM_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (mCARD_VALUE == "REPRE_NUM")
            {
                iedIC_CARD_NO.EditValue = e.EditValue;
            }
        }

        private void iedORI_JOIN_DATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (iString.ISNull(iedJOIN_DATE.EditValue) == string.Empty)
            {
                iedJOIN_DATE.EditValue = iedORI_JOIN_DATE.EditValue;
            }
            if (iString.ISNull(iedPAY_DATE.EditValue) == string.Empty)
            {
                iedPAY_DATE.EditValue = iedORI_JOIN_DATE.EditValue;
            }
        }

        private void iedJOIN_DATE_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            IDC_GET_PAY_DATE_P.ExecuteNonQuery(); 
            iedPAY_DATE.EditValue = IDC_GET_PAY_DATE_P.GetCommandParamValue("O_PAY_DATE"); 
        }

        private void igrFAMILY_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == igrFAMILY.GetColumnToIndex("REPRE_NUM"))
            {
                object vRepre_Num;
                vRepre_Num = e.NewValue;
                if (iString.ISNull(vRepre_Num) == string.Empty)
                {
                    return;
                }
                if (FAMILY_REPRE_NUM_CHECK(vRepre_Num) == "N".ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10026"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (iString.ISNull(igrFAMILY.GetCellValue("BIRTHDAY")) == string.Empty)
                {
                    igrFAMILY.SetCellValue("BIRTHDAY", BIRTHDAY(vRepre_Num));
                }

                if (iString.ISNull(igrFAMILY.GetCellValue("BIRTHDAY_TYPE")) == string.Empty)
                {
                    // 음양구분.
                    idcCOMMON_W.SetCommandParamValue("W_GROUP_CODE", "BIRTHDAY_TYPE");
                    idcCOMMON_W.SetCommandParamValue("W_WHERE", " 1 = 1 ");
                    idcCOMMON_W.ExecuteNonQuery();
                    igrFAMILY.SetCellValue("BIRTHDAY_TYPE_NAME", idcCOMMON_W.GetCommandParamValue("O_CODE_NAME"));
                    igrFAMILY.SetCellValue("BIRTHDAY_TYPE", idcCOMMON_W.GetCommandParamValue("O_CODE"));
                }
            }
        }

        private void iedNAME_0_KeyUp(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        private void iedLEGAL_ZIP_CODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_Legal();
            }
        }

        private void iedLEGAL_ADDR1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_Legal();
            }
        }

        private void iedPRSN_ZIP_CODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_PRSN();
            }
        }

        private void iedPRSN_ADDR1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_PRSN();
            }
        }

        private void iedLIVE_ZIP_CODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_Live();
            }
        }

        private void iedLIVE_ADDR1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_Live();
            }
        }

        private void igrCAREER_CellKeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                int vIDX_ROW = igrCAREER.RowIndex;
                int vIDX_ZIP_CODE = igrCAREER.GetColumnToIndex("ZIP_CODE");
                int vIDX_ADDR_1 = igrCAREER.GetColumnToIndex("ADDR1");
                if (igrCAREER.ColIndex == vIDX_ZIP_CODE || igrCAREER.ColIndex == vIDX_ADDR_1)
                {
                    Show_Address_Career(vIDX_ROW, vIDX_ZIP_CODE, vIDX_ADDR_1);
                }
            }
        }

        private void iedR_GUAR_ZIP_CODE1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_GUAR1();
            }
        }

        private void iedR_GUAR_ADDR1_1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_GUAR1();
            }
        }

        private void iedR_GUAR_ZIP_CODE2_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_GUAR2();
            }
        }

        private void iedR_GUAR_ADDR2_1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address_GUAR2();
            }
        }

        private void igrEDUCATION_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            int vIDX_START_DATE = igrEDUCATION.GetColumnToIndex("START_DATE");
            int vIDX_END_DATE = igrEDUCATION.GetColumnToIndex("END_DATE");
            object vSTART_DATE;
            object vEND_DATE;
            if(vIDX_START_DATE  == e.ColIndex)
            {
                vSTART_DATE = e.NewValue;
            }
            else
            {
                vSTART_DATE = igrEDUCATION.GetCellValue("START_DATE");
            }

            if(vIDX_END_DATE == e.ColIndex)
            {
                vEND_DATE = e.NewValue;
            }
            else 
            {
                vEND_DATE = igrEDUCATION.GetCellValue("END_DATE");
            }

            IDC_EDU_GET_TIME_P.SetCommandParamValue("P_START_DATE", vSTART_DATE);
            IDC_EDU_GET_TIME_P.SetCommandParamValue("P_END_DATE", vEND_DATE);
            IDC_EDU_GET_TIME_P.ExecuteNonQuery();
            igrEDUCATION.SetCellValue("EDU_TIME", IDC_EDU_GET_TIME_P.GetCommandParamValue("O_EDU_TIME")); 
        }

        private void BTN_IC_CARD_NUM_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(iString.ISNull(PERSON_ID.EditValue) == String.Empty)
            {
                return;
            }
            IGR_PERSON_IC_CARD.LastConfirmChanges();
            IDA_PERSON_IC_CARD.OraSelectData.AcceptChanges();
            IDA_PERSON_IC_CARD.Refillable = true;

            Sub_Form_Visible(true, "IC_CARD");
            IDA_PERSON_IC_CARD.Fill();
        }

        private void BTN_IC_INSERT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_PERSON_IC_CARD.AddUnder();
            IGR_PERSON_IC_CARD.Focus();
        }

        private void BTN_IC_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_PERSON_IC_CARD.Delete();
        }

        private void BTN_IC_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_PERSON_IC_CARD.Cancel();
        }

        private void BTN_IC_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Sub_Form_Visible(false, "IC_CARD");
        }

        private void BTN_IC_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_PERSON_IC_CARD.Update();
        }

        #endregion

        #region ----- Adapter Event -----

        private bool isDelete_Validate(string pTabPage)
        {
            bool ibReturn_Value = false;
            if (pTabPage == "itpPERSON_MASTER")
            {
                ibReturn_Value = false;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);   // 사원정보 삭제 불가.
            }
            return ibReturn_Value;
        }

// 인사기본 검증---------------------------------------------------------------
        private void idaPERSON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {// Added 상태가 아닐경우 체크.
                if (e.Row["PERSON_ID"] == DBNull.Value)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (string.IsNullOrEmpty(e.Row["PERSON_NUM"].ToString()))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            if (string.IsNullOrEmpty(e.Row["NAME"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Name(성명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["OPERATING_UNIT_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Operating Unit(사업장)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["DEPT_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Department(부서)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["NATION_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=국가"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["JOB_CLASS_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job Class(직군)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (e.Row["JOB_ID"] == DBNull.Value)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job(직종)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (e.Row["POST_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Post(직위)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["OCPT_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Ocpt(직무)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ABIL_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Abil(직책)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PAY_GRADE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Grade(직급)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            //if (string.IsNullOrEmpty(e.Row["REPRE_NUM"].ToString()))
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Repre Num(주민번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //    return;
            //}
            if (string.IsNullOrEmpty(e.Row["SEX_TYPE"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Sex Type(성별)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["JOIN_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=입사구분"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ORI_JOIN_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Ori Join Date(그룹입사일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["JOIN_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Join Date(입사일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["RETIRE_DATE"]) != string.Empty && iString.ISNull(e.Row["RETIRE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10170"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["RETIRE_DATE"]) == string.Empty && iString.ISNull(e.Row["RETIRE_ID"]) != string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10171"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["DIR_INDIR_TYPE"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Dir/InDir Type(직간접 구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }             
            if (e.Row["JOB_CATEGORY_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Job Category(직구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["FLOOR_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Floor(작업장)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaPERSON_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Person Infomation(인사정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 신체사항 검증---------------------------------------------------------------
        private void idaBODY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }

        private void idaBODY_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }


// 병역사항 검증---------------------------------------------------------------
        private void idaARMY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ARMY_KIND_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Army Kind(군별)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ARMY_STATUS_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Army Status(역종)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ARMY_GRADE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Army Grade(계급)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaARMY_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 가족사항 검증---------------------------------------------------------------
        private void idaFAMILY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["FAMILY_NAME"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Family Name(성명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["RELATION_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Family Relation(가족 관계)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["REPRE_NUM"]) != string.Empty)
            {
                if (FAMILY_REPRE_NUM_CHECK(e.Row["REPRE_NUM"]) == "N".ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10026"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void idaFAMILY_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 경력사항 검증---------------------------------------------------------------
        private void idaCAREE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["COMPANY_NAME"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation(회사명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (String.IsNullOrEmpty(e.Row["DEPT_NAME"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Department(부서명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["START_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Join Date(입사일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["END_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Retire Date(퇴사일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaCAREE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 학력사항 검증---------------------------------------------------------------
        private void idaSCHOLARSHIP_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["SCHOLARSHIP_TYPE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Scholarship Type(학력타입)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["GRADUATION_TYPE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Graduation Type(졸업구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ADMISSION_YYYYMM"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Admission Date(입학일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
            if (string.IsNullOrEmpty(e.Row["SCHOOL_NAME"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=School Name(학교명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaSCHOLARSHIP_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 교육사항 검증---------------------------------------------------------------
        private void idaEDUCATION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["START_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Education Start Date(교육 시작일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["EDU_CURRICULUM"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Education Curriculum(교육명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaEDUCATION_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 평가사항 검증---------------------------------------------------------------
        private void idaRESULT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["RESULT_YYYY"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Result Year Month(평가 년월)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaRESULT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        
// 자격사항 검증---------------------------------------------------------------
        private void idaLICENSE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["LICENSE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=License Kind(자격증 종류)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["LICENSE_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=License Get Date(취득일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaLICENSE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 어학사항 검증---------------------------------------------------------------
        private void idaFOREIGN_LANGUAGE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["EXAM_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Exam Date(응시 일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["EXAM_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Exam Kind(검정 종류)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaFOREIGN_LANGUAGE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 상벌사항 검증---------------------------------------------------------------
        private void idaREWARD_PUNISHMENT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["RP_TYPE"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Reward/Punishment Type(상벌구분)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["RP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Reward/Punishment Kind(상벌 항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["RP_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Reward/Punishment Date(상벌 일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["RP_ORG"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Reward/Punishment Organization(시행처"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaREWARD_PUNISHMENT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

// 신원보증 검증---------------------------------------------------------------
        private void idaREFERENCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {// 신원보증
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person Infomation(사원정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["REFERENCE_TYPE"].ToString() == "I".ToString())
            {
                if (string.IsNullOrEmpty(e.Row["INSUR_NAME"].ToString()))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Insurance Name(보험명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (e.Row["INSUR_START_DATE"] == DBNull.Value)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Insurance Start Date(보험시작일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            else if (e.Row["REFERENCE_TYPE"].ToString() == "P".ToString())
            {
                if (string.IsNullOrEmpty(e.Row["GUAR_NAME1"].ToString()))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Reference Name(보증인)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (string.IsNullOrEmpty(e.Row["GUAR_REPRE_NUM1"].ToString()))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Repre Num(주민번호)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                if (e.Row["GUAR_RELATION_ID1"] == DBNull.Value)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Reference Relation(보증인 관계)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
            else
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Reference Kind(보증유형)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            
        }

        private void idaREFERENCE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added && e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

        #region ----- idaPERSON NewRowMoved Event -----

        private void idaPERSON_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                ipbPERSON.ImageLocation = string.Empty;
                return;
            }

            DOC_ATT_FLAG(pBindingManager.DataRow["PERSON_NUM"]);  

            if (mIsFormLoad == true)
            {
                return;
            }
            isViewItemImage();
        }

        #endregion

        #region ----- lookup adapter event -----

        private void ilaYEAR_STR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            string Start_YEAR = "2000";
            string End_YEAR = DateTime.Today.AddYears(1).Year.ToString();

            ildYEAR_STR.SetLookupParamValue("W_START_YEAR", Start_YEAR);
            ildYEAR_STR.SetLookupParamValue("W_END_YEAR", End_YEAR);
        }

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_CORP_TYPE", "A");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaCORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaOPERATING_UNIT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_ID", W_CORP_ID.EditValue);
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (W_CORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            ildDEPT.SetLookupParamValue("W_CORP_ID", W_CORP_ID.EditValue);
            ildDEPT.SetLookupParamValue("W_DEPT_LEVEL", DBNull.Value);
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaEMPLOYE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("EMPLOYE_TYPE", null, "Y");
        }

        private void ILA_CONTRACT_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("CONTRACT_TYPE", null, "Y");
        }
        
        private void ILA_CONTRACT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("CONTRACT_TYPE", null, "Y");
        }

        private void ilaEMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("EMPLOYE_TYPE", null, "Y");
        }

        private void ilaOPERATING_UNIT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (iedCORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10011"), "Warning", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            ildOPERATING_UNIT.SetLookupParamValue("W_CORP_ID", iedCORP_ID.EditValue);
            ildOPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (iedCORP_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            ildDEPT.SetLookupParamValue("W_CORP_ID", iedCORP_ID.EditValue);
            ildDEPT.SetLookupParamValue("W_DEPT_LEVEL", DBNull.Value);
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaARMY_KIND_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("ARMY_KIND", null, "Y");
        }

        private void ilaARMY_STATUS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("ARMY_STATUS", null, "Y");
        }

        private void ilaARMY_GRADE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("ARMY_GRADE", null, "Y");
        }

        private void ilaARMY_END_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("ARMY_END_TYPE", null, "Y");
        }

        private void ilaEXCEPTION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("EXCEPTION", null, "Y");
        }

        private void ilaEXCEPTION_LICENSE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("LICENSE", null, "Y");
        }

        private void ilaEXCEPTION_GRADE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("LICENSE_GRADE", null, "Y");
        }
                
        private void ilaDIR_INDIR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("DIR_INDIR_TYPE", null, "Y");
        }

        private void ilaOCPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("OCPT", null, "Y");
        }

        private void ilaJOB_CLASS_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("JOB_CLASS", null, "Y");
        }

        private void ilaJOB_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("JOB", null, "Y");
        }

        private void ilaABIL_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("ABIL", null, "Y");
        }

        private void ilaPOST_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("POST", null, "Y");
        }

        private void ilaPAY_GRADE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("PAY_GRADE", null, "Y");
        }

        private void ilaBIRTHDAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("BIRTHDAY_TYPE", null, "Y");
        }

        private void ilaNATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("NATION", null, "Y");
        }

        private void ilaWORK_AREA_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("WORK_AREA", null, "Y");
        }

        private void ilaEND_SCH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("END_SCH", null, "Y");
        }

        private void ilaJOIN_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("JOIN", null, "Y");
        }

        private void ilaJOIN_ROUTE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("JOIN_ROUTE", null, "Y");
        }

        private void ilaRETIRE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("RETIRE", null, "Y");
        }

        private void ilaJOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("JOB_CATEGORY", null, "Y");
        }

        private void ilaWORK_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("WORK_TYPE", null, "Y");
        }

        private void ILA_FLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("FLOOR", null, "Y");
        }

        private void ilaFLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("FLOOR", null, "Y");
        }

        private void ilaRELIGION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("RELIGION", null, "Y");
        }

        private void ilaBLOOD_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("BLOOD_TYPE", null, "Y");
        }

        private void ilaACHRO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("ACHRO", null, "Y");
        }

        private void ilaDISABLED_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("DISABLED", null, "Y");
        }

        private void ilaBOHUN_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("BOHUN", null, "Y");
        }

        private void ilaBOHUN_RELATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("BOHUN_RELATION", null, "Y");
        }

        private void ilaF_RELATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("RELATION", null, "Y");
        }

        private void ilaF_BIRTHDAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("BIRTHDAY_TYPE", null, "Y");
        }

        private void ilaF_END_SCH_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("END_SCH", null, "Y");
        }

        private void ilaF_YEAR_DISABILITY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("YEAR_DISABILITY", null, "Y");
        }

        private void ilaSCHOLARSHIP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("SCHOLARSHIP_TYPE", null, "Y");
        }

        private void ilaGRADUATION_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("GRADUATION_TYPE", null, "Y");
        }

        private void ilaDEGREE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("DEGREE", null, "Y");
        }

        private void ilaLICENSE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("LICENSE", null, "Y");
        }

        private void ilaLICENSE_GRADE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("LICENSE_GRADE", null, "Y");
        }

        private void ilaEXAM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("EXAM", null, "Y");
        }

        private void ilaLANGUAGE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("LANGUAGE", null, "Y");
        }

        private void ILA_VALUER_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_TEAM_LEADER.SetLookupParamValue("W_YEAR_YYYY", igrRESULT.GetCellValue("RESULT_YYYY"));
        }

        private void ILA_VALUER_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_TEAM_LEADER.SetLookupParamValue("W_YEAR_YYYY", igrRESULT.GetCellValue("RESULT_YYYY"));
        }

        private void ILA_VALUER_3_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_TEAM_LEADER.SetLookupParamValue("W_YEAR_YYYY", igrRESULT.GetCellValue("RESULT_YYYY"));
        }

        private void ILA_VALUER_4_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_TEAM_LEADER.SetLookupParamValue("W_YEAR_YYYY", igrRESULT.GetCellValue("RESULT_YYYY"));
        }

        private void ILA_VALUER_5_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON_TEAM_LEADER.SetLookupParamValue("W_YEAR_YYYY", igrRESULT.GetCellValue("RESULT_YYYY"));
        }

        private void ilaRP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("RP_TYPE", null, "Y");
        }

        private void ilaRP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            string W_WHERE = Convert.ToString(igrREWARD_PUNISHMENT.GetCellValue("RP_TYPE"));
            W_WHERE = String.Format("{0}{1}{2}", "HC.VALUE1 = '", W_WHERE, "' ");
            ildCOMMON_W.SetLookupParamValue("W_GROUP_CODE", "RP");
            ildCOMMON_W.SetLookupParamValue("W_WHERE", W_WHERE);
            ildCOMMON_W.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaREFERENCE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("REFERENCE_TYPE", null, "Y");
        }

        private void ilaGUAR_RELATION_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("RELATION", null, "Y");
        }

        private void ilaGUAR_RELATION_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("RELATION", null, "Y");
        }

        private void ilaADMISSION_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Add(DateTime.Today, 1000)));
        }

        private void ilaGRADUATION_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Add(DateTime.Today, 1000)));
        }

        private void ilaCOST_CENTER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOST_CENTER.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_POST_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            isSetCommonLookUpParameter("POST", null, "Y");
        }

        private void ilaFLOOR_SelectedRowData(object pSender)
        {
            IDC_GET_FLOOR_REF_P.SetCommandParamValue("W_FLOOR_ID", iedFLOOR_ID.EditValue);
            IDC_GET_FLOOR_REF_P.ExecuteNonQuery();

            iedCOST_CENTER.EditValue = IDC_GET_FLOOR_REF_P.GetCommandParamValue("O_COST_CENTER_DESC");
            iedCOST_CENTER_ID.EditValue = IDC_GET_FLOOR_REF_P.GetCommandParamValue("O_COST_CENTER_ID");
            iedDIR_INDIR_TYPE.EditValue = IDC_GET_FLOOR_REF_P.GetCommandParamValue("O_DIR_INDIR_TYPE");
            iedDIR_INDIR_TYPE_NAME.EditValue = IDC_GET_FLOOR_REF_P.GetCommandParamValue("O_DIR_INDIR_TYPE_NAME");
        }

        #endregion

        #region ----- is View Item Image Method -----

        private void isViewItemImage()
        { 
            ipbPERSON.ImageLocation = string.Empty;
            ipbPERSON.ImageLocation = string.Format("{0}{1}.JPG", mPerson_ImageLocation, PERSON_NUM.EditValue);
            return; 
        }

        #endregion;

        #region ----- Make Directory ----

        private void MakeDirectory(string pFTP_Type)
        {
            if (pFTP_Type.Equals("PERSON_DOC"))
            {
                System.IO.DirectoryInfo vClient_Directory = new System.IO.DirectoryInfo(mClient_DocDirectory);
                if (vClient_Directory.Exists == false) //있으면 True, 없으면 False
                {
                    vClient_Directory.Create();
                }
            }
            else
            {
                System.IO.DirectoryInfo vClient_Directory = new System.IO.DirectoryInfo(mClient_ImageDirectory);
                if (vClient_Directory.Exists == false) //있으면 True, 없으면 False
                {
                    vClient_Directory.Create();
                }
            }
        }

        #endregion;

        #region ----- Image View ----

        private bool ImageView(string pFileName)
        {
            bool isView = false;

            bool isExist = System.IO.File.Exists(pFileName);
            if (isExist == true)
            {
                ipbPERSON.ImageLocation = pFileName;
                isView = true;
            }
            else
            {
                ipbPERSON.ImageLocation = string.Empty;
                isView = true;
            }
            return isView;
        }

        #endregion;

        #region ----- Get Information FTP Methods -----

        private bool GetInfomationFTP()
        {
            //사원사진 로케이션//
            mPerson_ImageLocation = "";
            try
            {
                idcFTP_INFO.SetCommandParamValue("W_FTP_CODE", "PERSON_PIC_VIEW");
                idcFTP_INFO.ExecuteNonQuery();

                mPerson_ImageLocation = string.Format("http://{0}:{1}{2}", idcFTP_INFO.GetCommandParamValue("O_HOST_IP")
                                                                    , idcFTP_INFO.GetCommandParamValue("O_HOST_PORT")
                                                                    , idcFTP_INFO.GetCommandParamValue("O_HOST_FOLDER")); 
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }

            bool isGet = false;
            try
            {
                idcFTP_INFO.SetCommandParamValue("W_FTP_CODE", "PERSON_PIC");
                idcFTP_INFO.ExecuteNonQuery();
                mImageFTP = new ItemImageInfomationFTP();

                mImageFTP.Host = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_HOST_IP"));
                mImageFTP.Port = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_HOST_PORT"));
                mImageFTP.UserID = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_USER_NO"));
                mImageFTP.Password = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_USER_PWD"));
                mImageFTP.Passive_Flag = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_PASSIVE_FLAG"));

                mFTP_Source_Directory = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_HOST_FOLDER"));
                mClient_Directory = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_CLIENT_FOLDER")); 

                mClient_ImageDirectory = string.Format("{0}\\{1}", mClient_Base_Path, mClient_Directory);

                if (mImageFTP.Host != string.Empty)
                {
                    isGet = true;
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
            return isGet;
        }

        private bool GetPersonDocFTP()
        {
            bool isGet = false;
            try
            {
                idcFTP_INFO.SetCommandParamValue("W_FTP_CODE", "PERSON_DOC");
                idcFTP_INFO.ExecuteNonQuery();
                mDocFTP = new PersonDocFTP();

                mDocFTP.Host = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_HOST_IP"));
                mDocFTP.Port = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_HOST_PORT"));
                mDocFTP.UserID = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_USER_NO"));
                mDocFTP.Password = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_USER_PWD"));
                mDocFTP.Passive_Flag = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_PASSIVE_FLAG"));

                mDocFTP.FTP_Folder = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_HOST_FOLDER"));
                mDocFTP.Client_Folder = iString.ISNull(idcFTP_INFO.GetCommandParamValue("O_CLIENT_FOLDER"));
                mClient_DocDirectory = string.Format("{0}\\{1}", mClient_Base_Path, mDocFTP.Client_Folder);

                if (mDocFTP.Host != string.Empty)
                {
                    isGet = true;
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
            return isGet;
        }

        #endregion;

        #region ----- FTP Initialize -----

        private void FTPInitializtion(string pFTP_Type)
        {            
            if (pFTP_Type.Equals("PERSON_DOC"))
            {
                mFileTransferAdv = new ISFileTransferAdv();
                mFileTransferAdv.Host = mDocFTP.Host; 
                mFileTransferAdv.Port = mDocFTP.Port;
                mFileTransferAdv.UserId = mDocFTP.UserID;
                mFileTransferAdv.Password = mDocFTP.Password;
                mFileTransferAdv.KeepAlive = false; 
                if (mDocFTP.Passive_Flag == "Y")
                {
                    mFileTransferAdv.UsePassive = true;
                }
                else
                {
                    mFileTransferAdv.UsePassive = false;
                }
            }
            else
            {
                mFileTransferAdv = new ISFileTransferAdv();
                mFileTransferAdv.Host = mImageFTP.Host;
                mFileTransferAdv.Port = mImageFTP.Port;
                mFileTransferAdv.UserId = mImageFTP.UserID;
                mFileTransferAdv.Password = mImageFTP.Password;
                mFileTransferAdv.KeepAlive = false;
                if (mImageFTP.Passive_Flag == "Y")
                {
                    mFileTransferAdv.UsePassive = true;
                }
                else
                {
                    mFileTransferAdv.UsePassive = false;
                }
            }
        }

        #endregion;

        #region ----- Image Upload Methods -----

        private bool UpLoadItem(string pPERSON_NUM)
        {
            bool isUp = false;

            openFileDialog1.FileName = string.Format("*{0}", mFileExtension);
            openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", mFileExtension, mFileExtension);
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string vChoiceFileFullPath = openFileDialog1.FileName;
                    string vChoiceFilePath = vChoiceFileFullPath.Substring(0, vChoiceFileFullPath.LastIndexOf(@"\"));
                    string vChoiceFileName = vChoiceFileFullPath.Substring(vChoiceFileFullPath.LastIndexOf(@"\") + 1);

                    mFileTransferAdv.ShowProgress = true;
                    //--------------------------------------------------------------------------------

                    string vSourceFileName = vChoiceFileName;

                    string vTargetFileName = string.Format("{0}{1}", pPERSON_NUM.ToUpper(), mFileExtension);

                    mFileTransferAdv.SourceDirectory = vChoiceFilePath;
                    mFileTransferAdv.SourceFileName = vSourceFileName;
                    mFileTransferAdv.TargetDirectory = mFTP_Source_Directory;
                    mFileTransferAdv.TargetFileName = vTargetFileName;

                    bool isUpLoad = mFileTransferAdv.Upload();

                    if (isUpLoad == true)
                    {
                        isUp = true;
                        bool isView = ImageView(vChoiceFileFullPath);
                    }
                    else
                    {
                    }
                }
                catch
                {
                }
            }
            System.IO.Directory.SetCurrentDirectory(mClient_Base_Path);
            return isUp;
        }

        private bool Delete_Item(string pPERSON_NUM)
        {
            bool isDel = false;
            string vTargetFileName = string.Format("{0}{1}", pPERSON_NUM.ToUpper(), mFileExtension);

            //Local 파일 삭제//
            try
            {
                string vFileName = string.Format("{0}\\{1}", mClient_ImageDirectory, vTargetFileName);
                if (System.IO.File.Exists(vFileName))
                {
                    System.IO.File.Delete(vFileName);
                }
            }
            catch
            {

            }

            //ftp server 삭제// 
            try
            {               
                mFileTransferAdv.ShowProgress = true;
                //--------------------------------------------------------------------------------

                mFileTransferAdv.SourceDirectory = mFTP_Source_Directory;  //삭제는 소스에 설정해야 삭제됨.
                mFileTransferAdv.SourceFileName = vTargetFileName;
                mFileTransferAdv.TargetDirectory = mFTP_Source_Directory;
                mFileTransferAdv.TargetFileName = vTargetFileName;

                bool isDelete = mFileTransferAdv.Delete();

                if (isDelete == true)
                {
                    isDel = true;
                    bool isView = ImageView("");
                } 
            }
            catch
            {
            }
            return isDel;
        }

        private void ibtPERSON_PICTURE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string vPerson_Num = iString.ISNull(PERSON_NUM.EditValue);
            if (vPerson_Num == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (mIsGetInformationFTP == true)
            {
                bool vResult = UpLoadItem(vPerson_Num);
                if (vResult == true)
                {
                    Save_Pic_Attach_File(PERSON_ID.EditValue, vPerson_Num, "Save");
                }
            }
        }

        private void BTN_DEL_PHOTO_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            string vPerson_Num = iString.ISNull(PERSON_NUM.EditValue);
            if (vPerson_Num == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (mIsGetInformationFTP == true)
            {
                bool vResult = Delete_Item(vPerson_Num);
                if (vResult == true)
                {
                    Save_Pic_Attach_File(PERSON_ID.EditValue, vPerson_Num, "DEL");
                }
            }
        }

        private void Save_Pic_Attach_File(object pPerson_ID, object pPerson_Num, string pSave_Type)
        {
            IDC_SAVE_PERSON_PIC_P.SetCommandParamValue("P_PERSON_ID", pPerson_ID);
            IDC_SAVE_PERSON_PIC_P.SetCommandParamValue("P_PERSON_NUM", pPerson_Num);
            IDC_SAVE_PERSON_PIC_P.SetCommandParamValue("P_SAVE_TYPE", pSave_Type);
            IDC_SAVE_PERSON_PIC_P.ExecuteNonQuery();
            string vStatus = iString.ISNull(IDC_SAVE_PERSON_PIC_P.GetCommandParamValue("O_STATUS"));
            string vMessage = iString.ISNull(IDC_SAVE_PERSON_PIC_P.GetCommandParamValue("O_MESSAGE"));
            if (vStatus == "F")
            {
                if (vMessage != string.Empty)
                {
                    MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
        }

        #endregion;

        #region ----- Image Download Methods -----

        private bool DownLoadItem(string pFileName)
        {
            bool isDown = false;

            string vSourceDownLoadFile = string.Format("{0}\\{1}", mClient_ImageDirectory, pFileName);
            string vTargetDownLoadFile = string.Format("{0}\\_{1}", mClient_ImageDirectory, pFileName);

            string vBeforeSourceFileName = string.Format("{0}", pFileName);
            string vBeforeTargetFileName = string.Format("_{0}", pFileName);

            mFileTransferAdv.ShowProgress = false;
            //-------------------------------------------------------------------------------- 
            mFileTransferAdv.SourceDirectory = mFTP_Source_Directory;
            mFileTransferAdv.SourceFileName = vBeforeSourceFileName;
            mFileTransferAdv.TargetDirectory = mClient_ImageDirectory;
            mFileTransferAdv.TargetFileName = vBeforeTargetFileName;
            
            isDown = mFileTransferAdv.Download();
            if (isDown == true)
            {
                try
                {
                    System.IO.File.Delete(vSourceDownLoadFile);
                    System.IO.File.Move(vTargetDownLoadFile, vSourceDownLoadFile);

                    isDown = true;
                }
                catch
                {
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTargetDownLoadFile);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTargetDownLoadFile);
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    }
                    catch
                    {
                        //ignore
                    }
                }
            }
            else
            {
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTargetDownLoadFile);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTargetDownLoadFile);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                catch
                {
                    //ignore
                }
            }

            return isDown;
        }

        #endregion;


        #region ----- File Upload Methods -----
        //ftp에 file upload 처리 
        private bool UpLoadFile(object pPERSON_NUM)
        {
            bool isUpload = false;
            OpenFileDialog vOpenFileDialog1 = new OpenFileDialog();
            vOpenFileDialog1.RestoreDirectory = true;

            if (!mIsGetPersonDocFTP)
            {
                isAppInterfaceAdv1.OnAppMessage("FTP Server Connect Fail. Check FTP Server");
                return isUpload;
            }

            if (iString.ISNull(pPERSON_NUM) != string.Empty)
            {
                string vSTATUS = "F";
                string vMESSAGE = string.Empty;

                //openFileDialog1.FileName = string.Format("*{0}", vFileExtension);
                //openFileDialog1.Filter = string.Format("Image Files (*{0})|*{1}", vFileExtension, vFileExtension);

                vOpenFileDialog1.Title = "Select Open File";
                vOpenFileDialog1.Filter = "All File(*.*)|*.*|pdf File(*.pdf)|*.pdf|jpg file(*.jpg)|*.jpg|bmp file(*.bmp)|*.bmp";
                vOpenFileDialog1.DefaultExt = "*.*";
                vOpenFileDialog1.FileName = "";
                vOpenFileDialog1.RestoreDirectory = true;
                vOpenFileDialog1.Multiselect = true;


                if (vOpenFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Application.UseWaitCursor = true;
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();

                    string vSelectFullPath = string.Empty;
                    string vSelectDirectoryPath = string.Empty;

                    string vFileName = string.Empty;
                    string vFileExtension = string.Empty;

                    //1. 사용자 선택 파일 
                    for (int i = 0; i < vOpenFileDialog1.FileNames.Length; i++)
                    {
                        vSelectFullPath = vOpenFileDialog1.FileNames[i];
                        vSelectDirectoryPath = System.IO.Path.GetDirectoryName(vSelectFullPath);

                        vFileName = System.IO.Path.GetFileName(vSelectFullPath);
                        vFileExtension = System.IO.Path.GetExtension(vSelectFullPath).ToUpper();

                        //2. 첨부파일 DB 저장 
                        IDC_INSERT_DOC_ATTACH.SetCommandParamValue("P_SOURCE_CATEGORY", "PERSON_DOC"); //구분  
                        IDC_INSERT_DOC_ATTACH.SetCommandParamValue("P_SOURCE_NUM", pPERSON_NUM);
                        IDC_INSERT_DOC_ATTACH.SetCommandParamValue("P_USER_FILE_NAME", vFileName);
                        IDC_INSERT_DOC_ATTACH.SetCommandParamValue("P_FTP_FILE_NAME", vFileName);
                        IDC_INSERT_DOC_ATTACH.SetCommandParamValue("P_EXTENSION_NAME", vFileExtension);
                        IDC_INSERT_DOC_ATTACH.ExecuteNonQuery();

                        vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACH.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACH.GetCommandParamValue("O_MESSAGE"));
                        object vDOC_ATTACH_ID = IDC_INSERT_DOC_ATTACH.GetCommandParamValue("O_DOC_ATTACH_ID");
                        object vFTP_FILE_NAME = IDC_INSERT_DOC_ATTACH.GetCommandParamValue("O_FTP_FILE_NAME");

                        //O_DOC_ATTACHMENT_ID.EditValue = vDOC_ATTACHMENT_ID;
                        //O_FTP_FILE_NAME.EditValue = vFTP_FILE_NAME;

                        if (IDC_INSERT_DOC_ATTACH.ExcuteError || vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();

                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                MessageBoxAdv.Show(IDC_INSERT_DOC_ATTACH.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return isUpload;
                        }

                        //3. 첨부파일 로그 저장 
                        IDC_INSERT_DOC_ATTACH_LOG.SetCommandParamValue("P_DOC_ATTACH_ID", vDOC_ATTACH_ID);
                        IDC_INSERT_DOC_ATTACH_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "IN");
                        IDC_INSERT_DOC_ATTACH_LOG.ExecuteNonQuery();
                        vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACH_LOG.GetCommandParamValue("O_STATUS"));
                        vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACH_LOG.GetCommandParamValue("O_MESSAGE"));
                        if (IDC_INSERT_DOC_ATTACH_LOG.ExcuteError || vSTATUS == "F")
                        {
                            Application.UseWaitCursor = false;
                            this.Cursor = Cursors.Default;
                            Application.DoEvents();
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            return isUpload;
                        }

                        //4. 파일 업로드
                        try
                        { 
                            mFileTransferAdv.ShowProgress = true;      //진행바 보이기 

                            //업로드 환경 설정 
                            mFileTransferAdv.SourceDirectory = vSelectDirectoryPath;
                            mFileTransferAdv.SourceFileName = vFileName;
                            mFileTransferAdv.TargetDirectory = mDocFTP.FTP_Folder;
                            mFileTransferAdv.TargetFileName = iString.ISNull(vFTP_FILE_NAME);

                            bool isUpLoad = mFileTransferAdv.Upload();

                            if (isUpLoad == true)
                            {
                                isUpload = true;
                            }
                            else
                            {
                                isUpload = false;
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10092"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                            //5. 적용 
                        }
                        catch (Exception Ex)
                        {
                            isAppInterfaceAdv1.OnAppMessage(Ex.Message);
                            return isUpload;
                        }
                    }
                }
            }
            return isUpload;
        }

        #endregion;


        #region ----- file Download Methods -----
        //ftp file download 처리 
        private bool DownLoadFile(object pDOC_ATTACH_ID, string pFTP_FileName, string pClient_FileName)
        {
            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            bool IsDownload = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            ////1. 첨부파일 로그 저장 : Transaction을 이용해서 처리 
            //isDataTransaction1.BeginTran();            
            IDC_INSERT_DOC_ATTACH_LOG.SetCommandParamValue("P_DOC_ATTACH_ID", pDOC_ATTACH_ID);
            IDC_INSERT_DOC_ATTACH_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "OUT");
            IDC_INSERT_DOC_ATTACH_LOG.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACH_LOG.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACH_LOG.GetCommandParamValue("O_MESSAGE"));
            if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return IsDownload;
            }

            //2. 실제 다운로드 
            string vTempFileName = string.Format("_{0}", pFTP_FileName);
            try
            {
                System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                if (vDownFileInfo.Exists == true)
                {
                    try
                    {
                        System.IO.File.Delete(vTempFileName);
                    }
                    catch
                    {

                        // ignore
                    }
                }
            }
            catch
            {
                //ignore                        
            }

            mFileTransferAdv.ShowProgress = false;
            //--------------------------------------------------------------------------------
            mFileTransferAdv.SourceDirectory = mDocFTP.FTP_Folder;
            mFileTransferAdv.SourceFileName = pFTP_FileName;
            mFileTransferAdv.TargetDirectory = mClient_DocDirectory;
            mFileTransferAdv.TargetFileName = vTempFileName;

            IsDownload = mFileTransferAdv.Download();

            if (IsDownload == true)
            {
                try
                {
                    //isDataTransaction1.Commit();

                    //다운 파일 FullPath적용 
                    string vTempFullPath = string.Format("{0}\\{1}", mClient_DocDirectory, vTempFileName);      //임시

                    System.IO.File.Delete(pClient_FileName);                 //기존 파일 삭제 
                    System.IO.File.Move(vTempFullPath, pClient_FileName);    //ftp 이름으로 이름 변경 

                    IsDownload = true;
                }
                catch
                {
                    //isDataTransaction1.RollBack();
                    try
                    {
                        System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                        if (vDownFileInfo.Exists == true)
                        {
                            try
                            {
                                System.IO.File.Delete(vTempFileName);
                            }
                            catch
                            {

                                // ignore
                            }
                        }
                    }
                    catch
                    {
                        //ignore                        
                    }
                }
            }
            else
            {
                //isDataTransaction1.RollBack();
                //download 실패 
                try
                {
                    System.IO.FileInfo vDownFileInfo = new System.IO.FileInfo(vTempFileName);
                    if (vDownFileInfo.Exists == true)
                    {
                        try
                        {
                            System.IO.File.Delete(vTempFileName);
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                }
                catch
                {
                    //ignore                    
                }
            }
            if (IsDownload == true)
            {
                System.Diagnostics.Process.Start(pClient_FileName);
            }
            else
            {
                string vMessage = string.Format("{0} {1}", isMessageAdapter1.ReturnText("EAPP_10212"), isMessageAdapter1.ReturnText("QM_10102"));
                MessageBoxAdv.Show(vMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
            return IsDownload;
        }

        #endregion;

        #region ----- file Delete Methods -----
        //ftp file delete 처리 
        private bool DeleteFile(object pDOC_ATTACH_ID, string pFTP_FileName)
        {
            bool IsDelete = false;
            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            if (iString.ISNull(pDOC_ATTACH_ID) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }
            if (pFTP_FileName == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10075"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return IsDelete;
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();


            //1. 첨부파일 로그 저장 : Transaction을 이용해서 처리  
            IDC_INSERT_DOC_ATTACH_LOG.SetCommandParamValue("P_DOC_ATTACH_ID", pDOC_ATTACH_ID);
            IDC_INSERT_DOC_ATTACH_LOG.SetCommandParamValue("P_IN_OUT_STATUS", "DELETE");
            IDC_INSERT_DOC_ATTACH_LOG.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_INSERT_DOC_ATTACH_LOG.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_INSERT_DOC_ATTACH_LOG.GetCommandParamValue("O_MESSAGE"));
            if (IDC_INSERT_DOC_ATTACH_LOG.ExcuteError || vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                this.Cursor = Cursors.Default;
                Application.DoEvents();

                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBoxAdv.Show(IDC_INSERT_DOC_ATTACH_LOG.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return IsDelete;
            }

            //2. 파일 삭제 
            IDC_DELETE_DOC_ATTACH.SetCommandParamValue("W_DOC_ATTACH_ID", pDOC_ATTACH_ID);
            IDC_DELETE_DOC_ATTACH.ExecuteNonQuery();
            vSTATUS = iString.ISNull(IDC_DELETE_DOC_ATTACH.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iString.ISNull(IDC_DELETE_DOC_ATTACH.GetCommandParamValue("O_MESSAGE"));

            if (IDC_DELETE_DOC_ATTACH.ExcuteError || vSTATUS == "F")
            {
                IsDelete = false;
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();


                if (vMESSAGE != string.Empty)
                {
                    MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBoxAdv.Show(IDC_DELETE_DOC_ATTACH.ExcuteErrorMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return IsDelete;
            }

            //3. 실제 삭제 
            mFileTransferAdv.ShowProgress = false;
            //--------------------------------------------------------------------------------

            mFileTransferAdv.SourceDirectory = mDocFTP.FTP_Folder;
            mFileTransferAdv.SourceFileName = pFTP_FileName;
            mFileTransferAdv.TargetDirectory = mDocFTP.FTP_Folder;
            mFileTransferAdv.TargetFileName = pFTP_FileName;

            IsDelete = mFileTransferAdv.Delete();
            if (IsDelete == false)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                return IsDelete;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            return IsDelete;
        }

        #endregion; 


        private bool Check_Sub_Panel()
        {
            if (mSUB_SHOW_FLAG == true)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10069"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
         
        private void DOC_ATT_FLAG(object pPERSON_NUM)
        { 
            IDC_GET_DOC_ATT_FLAG_P.SetCommandParamValue("P_SOURCE_CATEGORY", "PERSON_DOC");
            IDC_GET_DOC_ATT_FLAG_P.SetCommandParamValue("P_SOURCE_NUM", pPERSON_NUM);
            IDC_GET_DOC_ATT_FLAG_P.ExecuteNonQuery();
            if (iString.ISNull(IDC_GET_DOC_ATT_FLAG_P.GetCommandParamValue("O_DOC_ATT_FLAG")) == "Y")
            {
                CB_DOC_ATT_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;
            }
            else
            {
                CB_DOC_ATT_FLAG.CheckedState = ISUtil.Enum.CheckedState.Unchecked;
            }
        }

        private void BTN_FILE_ATTACH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            F_PERSON_NUM.EditValue = PERSON_NUM.EditValue;  
            if (iString.ISNull(F_PERSON_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Sub_Form_Visible(true, "PERSON_DOC");

            //FTP 정보//
            FTPInitializtion("PERSON_DOC");

            IDA_DOC_ATTACH.Fill();
            IGR_DOC_ATTACH.Focus();
        }
         
        private void IGR_DOC_ATTACH_CellDoubleClick(object pSender)
        {
            //if (IGR_DOC_ATTACHMENT.RowIndex < 0)
            //{
            //    return;
            //}

            //string vFTP_FILE_NAME = iString.ISNull(IGR_DOC_ATTACHMENT.GetCellValue("FTP_FILE_NAME"));
            //string vUSER_FILE_NAME = string.Format("{0}{1}", mDownload_Folder, IGR_DOC_ATTACHMENT.GetCellValue("USER_FILE_NAME"));
            //if (DownLoadFile(vFTP_FILE_NAME, vUSER_FILE_NAME) == false)
            //{
            //    return;
            //} 
        }

        private void BTN_ATT_SELECT_ButtonClick(object pSender, EventArgs pEventArgs)
        { 
            UpLoadFile(F_PERSON_NUM.EditValue);
            IDA_DOC_ATTACH.Fill();
            IGR_DOC_ATTACH.Focus();

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();
        }

        private void BTN_ATT_DOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IGR_DOC_ATTACH.RowIndex < 0)
            {
                return;
            }

            object vDOC_ATTACH_ID = IGR_DOC_ATTACH.GetCellValue("DOC_ATTACH_ID");
            string vFTP_FILE_NAME = iString.ISNull(IGR_DOC_ATTACH.GetCellValue("FTP_FILE_NAME"));
            string vUSER_FILE_NAME = string.Format("{0}\\{1}", mClient_DocDirectory, IGR_DOC_ATTACH.GetCellValue("USER_FILE_NAME"));
            if (DownLoadFile(vDOC_ATTACH_ID, vFTP_FILE_NAME, vUSER_FILE_NAME) == false)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();
                return;
            }
        }

        private void BTN_ATT_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10220"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }
            if (iString.ISNull(F_PERSON_NUM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10218"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
             
            if (IGR_DOC_ATTACH.RowIndex < 0)
            {
                return;
            }

            object vDOC_ATTACH_ID = IGR_DOC_ATTACH.GetCellValue("DOC_ATTACH_ID");
            string vFTP_FileName = iString.ISNull(IGR_DOC_ATTACH.GetCellValue("FTP_FILE_NAME"));
            DeleteFile(vDOC_ATTACH_ID, vFTP_FileName);
            IDA_DOC_ATTACH.Fill();
            IGR_DOC_ATTACH.Focus();
        }

        private void BTN_ATT_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DOC_ATT_FLAG(PERSON_NUM.EditValue);
            Sub_Form_Visible(false, "PERSON_DOC");
        }

        private void GB_DOC_ATT_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = true;
            mInquiryDetailPreX = e.X;
            mInquiryDetailPreY = e.Y;
        }

        private void GB_DOC_ATT_MouseMove(object sender, MouseEventArgs e)
        {
            if (mIsClickInquiryDetail && e.Button == MouseButtons.Left)
            {
                int gx = e.X - mInquiryDetailPreX;
                int gy = e.Y - mInquiryDetailPreY;

                Point I = GB_DOC_ATT.Location;
                I.Offset(gx, gy);
                GB_DOC_ATT.Location = I;
            }
        }
         
        private void GB_DOC_ATT_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClickInquiryDetail = false;
        }


        //#region ----- Get Registry Customer Methods ----

        //private string GetRegistryCustomer()
        //{
        //    string vMessage = string.Empty;
        //    string vCustomer = "BH";

        //    //C:\Program Files\Flex_ERP_BH\Kor
        //    //string vWorkDirectory = "C:\\Program Files\\Flex_ERP_BH\\Kor";
        //    //string vWorkDirectory = "C:\\Program Files\\Flex_ERP_FC\\Kor";

        //    //-------------------------------------------------------------------
        //    string vWorkDirectory = System.Windows.Forms.Application.StartupPath;
        //    //-------------------------------------------------------------------

        //    int vCutLast1 = vWorkDirectory.LastIndexOf("\\");
        //    string vCutString = vWorkDirectory.Substring(0, vCutLast1);
        //    int vCutLast2 = vCutString.LastIndexOf("\\") + 1;
        //    int vLength = vCutLast1 - vCutLast2;
        //    string vCustomerString = vCutString.Substring(vCutLast2, vLength);

        //    string vFTPKey = string.Format(@"Software\{0}\{1}\{2}", "InfoSummit", vCustomerString, "FTP");
        //    Microsoft.Win32.RegistryKey vKey = Microsoft.Win32.Registry.LocalMachine;

        //    try
        //    {
        //        vKey = vKey.OpenSubKey(vFTPKey, true);

        //        vCustomer = vKey.GetValue("Customer").ToString();

        //        vKey.Close();
        //    }
        //    catch (System.Exception ex)
        //    {
        //        vMessage = ex.Message;
        //    }

        //    return vCustomer;
        //}

        //#endregion;

    }

    #region ----- User Make Class -----

    public class ItemImageInfomationFTP
    {
        #region ----- Variables -----

        private string mHost = string.Empty;
        private string mPort = "21";
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mPassive_Flag = "N";

        #endregion;

        #region ----- Constructor -----

        public ItemImageInfomationFTP()
        {
        }

        public ItemImageInfomationFTP(string pHost, string pPort, string pUserID, string pPassword, string pPassive_Flag)
        {
            mHost = pHost;
            mPort = pPort;
            mUserID = pUserID;
            mPassword = pPassword;
            mPassive_Flag = pPassive_Flag;
        }

        #endregion;

        #region ----- Property -----

        public string Host
        {
            get
            {
                return mHost;
            }
            set
            {
                mHost = value;
            }
        }

        public string Port
        {
            get
            {
                return mPort;
            }
            set
            {
                mPort = value;
            }
        }

        public string UserID
        {
            get
            {
                return mUserID;
            }
            set
            {
                mUserID = value;
            }
        }

        public string Password
        {
            get
            {
                return mPassword;
            }
            set
            {
                mPassword = value;
            }
        }

        public string Passive_Flag
        {
            get
            {
                return mPassive_Flag;
            }
            set
            {
                mPassive_Flag = value;
            }
        }

        #endregion;
    }

    #endregion;


    #region ----- User Make Person Doc Class -----

    public class PersonDocFTP
    {
        #region ----- Variables -----

        private string mHost = string.Empty;
        private string mPort = "21";
        private string mUserID = string.Empty;
        private string mPassword = string.Empty;
        private string mPassive_Flag = "N";
        private string mFTP_Folder = string.Empty;
        private string mClient_Folder = string.Empty;

        #endregion;

        #region ----- Constructor -----

        public PersonDocFTP()
        {
        }

        public PersonDocFTP(string pHost, string pPort, string pUserID, string pPassword, string pPassive_Flag)
        {
            mHost = pHost;
            mPort = pPort;
            mUserID = pUserID;
            mPassword = pPassword;
            mPassive_Flag = pPassive_Flag;
        }

        #endregion;

        #region ----- Property -----

        public string Host
        {
            get
            {
                return mHost;
            }
            set
            {
                mHost = value;
            }
        }

        public string Port
        {
            get
            {
                return mPort;
            }
            set
            {
                mPort = value;
            }
        }

        public string UserID
        {
            get
            {
                return mUserID;
            }
            set
            {
                mUserID = value;
            }
        }

        public string Password
        {
            get
            {
                return mPassword;
            }
            set
            {
                mPassword = value;
            }
        }

        public string Passive_Flag
        {
            get
            {
                return mPassive_Flag;
            }
            set
            {
                mPassive_Flag = value;
            }
        }

        public string FTP_Folder
        {
            get
            {
                return mFTP_Folder;
            }
            set
            {
                mFTP_Folder = value;
            }
        }

        public string Client_Folder
        {
            get
            {
                return mClient_Folder;
            }
            set
            {
                mClient_Folder = value;
            }
        }

        #endregion;
    }

    #endregion;
}