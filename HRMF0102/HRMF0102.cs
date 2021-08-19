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

namespace HRMF0102
{
    public partial class HRMF0102 : Office2007Form
    {
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public HRMF0102(Form pMainFom, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainFom;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Property Method ------

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void SEARCH_DB()
        {
            IDA_CORP_MASTER.Fill();
            IGR_CORP_LIST.Focus();
        }

        private void Insert_Corporation()
        {

            EFFECTIVE_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);
            ENABLED_FLAG.CheckBoxValue = "Y";

            CORP_NAME.Focus();
        }

        private void Insert_Operating_Unit()
        {
            igrHRM_OPERATING_UNIT_G.SetCellValue("ENABLED_FLAG", "Y");
            igrHRM_OPERATING_UNIT_G.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));

            igrHRM_OPERATING_UNIT_G.CurrentCellMoveTo(igrHRM_OPERATING_UNIT_G.GetColumnToIndex("OPERATING_UNIT_NAME"));
            igrHRM_OPERATING_UNIT_G.Focus();
        }
        #endregion

        #region ----- 주소관리 ----

        private void Show_Address()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, ZIP_CODE.EditValue, ADDR1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                ZIP_CODE.EditValue = vEAPF0299.Get_Zip_Code;
                ADDR1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void Show_Address_OPERATING_UNIT(int pIDX_Row, int pIDX_ZIP_CODE, int pIDX_ADDRESS)
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            igrHRM_OPERATING_UNIT_G.LastConfirmChanges();
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent
                                                    , isAppInterfaceAdv1.AppInterface
                                                    , igrHRM_OPERATING_UNIT_G.GetCellValue(pIDX_Row, pIDX_ZIP_CODE)
                                                    , igrHRM_OPERATING_UNIT_G.GetCellValue(pIDX_Row, pIDX_ADDRESS));
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                igrHRM_OPERATING_UNIT_G.SetCellValue(pIDX_Row, pIDX_ZIP_CODE, vEAPF0299.Get_Zip_Code);
                igrHRM_OPERATING_UNIT_G.SetCellValue(pIDX_Row, pIDX_ADDRESS, vEAPF0299.Get_Address);
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        #endregion

        #region ----- main Button Click -----
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
                    if (IDA_CORP_MASTER.IsFocused)
                    {
                        IDA_CORP_MASTER.AddOver();
                        Insert_Corporation();
                    }
                    else if (IDA_OPERATING_UNIT.IsFocused)
                    {
                        IDA_OPERATING_UNIT.AddOver();
                        Insert_Operating_Unit();

                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_CORP_MASTER.IsFocused)
                    {
                        IDA_CORP_MASTER.AddUnder();
                        Insert_Corporation();
                    }
                    else if (IDA_OPERATING_UNIT.IsFocused)
                    {
                        IDA_OPERATING_UNIT.AddUnder();
                        Insert_Operating_Unit();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_CORP_MASTER.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_CORP_MASTER.IsFocused)
                    {
                        IDA_CORP_MASTER.Cancel();
                    }
                    else if (IDA_OPERATING_UNIT.IsFocused)
                    {
                        IDA_OPERATING_UNIT.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_CORP_MASTER.IsFocused)
                    {
                        IDA_CORP_MASTER.Delete();
                    }
                    else if (IDA_OPERATING_UNIT.IsFocused)
                    {
                        IDA_OPERATING_UNIT.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----

        private void HRMF0102_Load(object sender, EventArgs e)
        {
            IDA_CORP_MASTER.FillSchema();

            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }

        private void iedZIP_CODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address();
            }
        }

        private void iedADDR1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address();
            }
        }

        private void igrHRM_OPERATING_UNIT_G_CellKeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                int vIDX_ROW = igrHRM_OPERATING_UNIT_G.RowIndex;
                int vIDX_ZIP_CODE = igrHRM_OPERATING_UNIT_G.GetColumnToIndex("ZIP_CODE");
                int vIDX_ADDR_1 = igrHRM_OPERATING_UNIT_G.GetColumnToIndex("ADDR1");
                if (igrHRM_OPERATING_UNIT_G.ColIndex == vIDX_ZIP_CODE || igrHRM_OPERATING_UNIT_G.ColIndex == vIDX_ADDR_1)
                {
                    Show_Address_OPERATING_UNIT(vIDX_ROW, vIDX_ZIP_CODE, vIDX_ADDR_1);
                }
            }
        }

        #endregion

        #region ----- ADAPTER EVENT ------

        private void IDA_CORP_MASTER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            // 필수 입력 데이터 검증 //
            if (string.IsNullOrEmpty(e.Row["CORP_NAME"].ToString()))
            {// 업체명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 업체정보
                CORP_NAME.Focus();
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["PRESIDENT_NAME"].ToString()))
            {// 대표자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10002"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 대표자 성명
                PRESIDENT_NAME.Focus();
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["CORP_CATEGORY_NAME"].ToString()))
            {//법인구분
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10003"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 법인구분
                CORP_CATEGORY_NAME.Focus();
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["LEGAL_NUMBER"].ToString()))
            {//법인번호
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10004"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 법인번호
                LEGAL_NUMBER.Focus();
                e.Cancel = true;
                return;
            } 
            if (e.Row["EFFECTIVE_DATE_FR"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자
                EFFECTIVE_DATE_FR.Focus();
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_TO"] != DBNull.Value)
            {
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자~종료일자
                    EFFECTIVE_DATE_FR.Focus();
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void IDA_CORP_MASTER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void IDA_OPERATING_UNIT_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Row["OPERATING_UNIT_NAME"].ToString()))
            {// 사업장명.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10007"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 사업장
                e.Cancel = true;
                return;
            }
            if (string.IsNullOrEmpty(e.Row["PRESIDENT_NAME"].ToString()))
            {// 대표자.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10002"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_FR"] == DBNull.Value)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_TO"] != DBNull.Value)
            {
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자~종료일자
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void IDA_OPERATING_UNIT_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
        #endregion        

        #region ----- LOOKUP EVENT -----

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CORP_ALL.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaCORP_TYPE1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "CORP_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_BIZ_UNIT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "BIZ_UNIT_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaCORP_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "CORP_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaADDRESS1_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildADDRESS1.SetLookupParamValue("W_ADDRESS", e.FilterString);
        }

        private void ilaADDRESS2_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ildADDRESS2.SetLookupParamValue("W_ADDRESS", e.FilterString);
        }
        
        private void ILA_VENDOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_VENDOR.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_TAX_OFFICE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "TAX_OFFICE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_BANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_BANK_ACCOUNT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_BANK_ACCOUNT.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ilaRETIRE_IRP_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "RETIRE_IRP_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

        #region ----- Convert decimal  Method ----

        private decimal ConvertNumber(object pObject)
        {
            bool vIsConvert = false;
            decimal vConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is decimal;
                    if (vIsConvert == true)
                    {
                        decimal vIsConvertNum = (decimal)pObject;
                        vConvertDecimal = vIsConvertNum;
                    }
                }

            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                //System.Windows.Forms.Application.DoEvents();
            }

            return vConvertDecimal;
        }

        #endregion;

    }
}