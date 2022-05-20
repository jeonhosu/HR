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

namespace HRMF2101
{
    public partial class HRMF2101 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public HRMF2101(Form pMainFom, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainFom;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Method -----
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

        }

        private void SEARCH_DB()
        {
            if (itbDEPT.SelectedTab.TabIndex == TP_DEPT_DETAIL.TabIndex)
            {
                IDA_DEPT_MASTER.Fill();
                IGR_DEPT_DETAIL.Focus();
            }
            else if (itbDEPT.SelectedTab.TabIndex == TP_DEPT_MAPPING.TabIndex)
            {
                if (iString.ISNull(W_MODULE_TYPE.EditValue) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10130"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    W_MODULE_TYPE_NAME.Focus();
                    return;
                }
                IDA_DEPT_MAPPING.Fill();
                IGR_DEPT_MAPPING.Focus();
            }
            else if (itbDEPT.SelectedTab.TabIndex == TP_DEPT_LIST.TabIndex)
            {
                IDA_DEPT_MASTER_LIST.Fill();
                IGR_DEPT_MASTER_LIST.Focus();
            }
        }

        private void Insert_Dept()
        {
            IGR_DEPT_DETAIL.SetCellValue("ENABLED_FLAG", "Y");
            IGR_DEPT_DETAIL.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
        }

        private void Insert_Dept_Mapping()
        {
            IGR_DEPT_MAPPING.SetCellValue("MODULE_TYPE", W_MODULE_TYPE.EditValue);
            IGR_DEPT_MAPPING.SetCellValue("MODULE_TYPE_NAME", W_MODULE_TYPE_NAME.EditValue); 
            IGR_DEPT_MAPPING.SetCellValue("ENABLED_FLAG", "Y");
            IGR_DEPT_MAPPING.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);

        }

        #endregion

        #region ----- main Button Click ------
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
                    if (IDA_DEPT_MASTER.IsFocused)
                    {
                        IDA_DEPT_MASTER.AddOver();
                        Insert_Dept();
                    }
                    else if (IDA_DEPT_MAPPING.IsFocused)
                    {
                        IDA_DEPT_MAPPING.AddOver();
                        Insert_Dept_Mapping();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_DEPT_MASTER.IsFocused)
                    {
                        IDA_DEPT_MASTER.AddUnder();
                        Insert_Dept();
                    }
                    else if (IDA_DEPT_MAPPING.IsFocused)
                    {
                        IDA_DEPT_MAPPING.AddUnder();
                        Insert_Dept_Mapping();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_DEPT_MASTER.IsFocused)
                    {
                       IDA_DEPT_MASTER.Update();
                    }
                    else if (IDA_DEPT_MAPPING.IsFocused)
                    {
                        IDA_DEPT_MAPPING.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DEPT_MASTER.IsFocused)
                    {
                        IDA_DEPT_MASTER.Cancel();
                    }
                    else if (IDA_DEPT_MAPPING.IsFocused)
                    {
                        IDA_DEPT_MAPPING.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DEPT_MASTER.IsFocused)
                    {
                        IDA_DEPT_MASTER.Delete();
                    }
                    else if (IDA_DEPT_MAPPING.IsFocused)
                    {
                        IDA_DEPT_MAPPING.Delete();
                    }
                }
            }
        }
        #endregion

        #region ----- Form Event -----
        private void HRMF2101_Load(object sender, EventArgs e)
        {
            IDA_DEPT_MASTER.FillSchema();
            IDA_DEPT_MAPPING.FillSchema();

            DefaultCorporation();
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }
        #endregion

        #region ---- Adapter Event -----
        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {             
            if (iString.ISNull(e.Row["DEPT_CODE"]) == string.Empty)
            {// 부서코드
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10019"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DEPT_NAME"]) == string.Empty)
            {// 부서명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10020"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DEPT_LEVEL"]) == string.Empty)
            {// 부서 레벨
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10021"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNumtoZero(e.Row["DEPT_LEVEL"]) > Convert.ToInt32(1))
            {// 부서 레벨이 0이 아닐경우 상위부서는 반드시 선택해야 합니다.
                if (iString.ISNull(e.Row["UPPER_DEPT_ID"]) == string.Empty)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10132"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            } 
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {// 시작일자 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_TO"]) != string.Empty)
            {// 종료일자 
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaDEPT_MAPPING_PreRowUpdate(ISPreRowUpdateEventArgs e)
        { 
            if (iString.ISNull(e.Row["MODULE_TYPE"]) == String.Empty)
            {// 모듈명
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10130"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["HR_DEPT_ID"]) == String.Empty)
            {// 인사부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10020"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["M_DEPT_ID"]) == String.Empty)
            {// 맵핑부서
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10131"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == String.Empty)
            {// 시작일자 
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_TO"]) != String.Empty)
            {// 종료일자 
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void idaDEPT_MAPPING_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion

        #region ---- Lookup Event -----

        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaDEPT_UPPER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT_UPPER.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaMODULE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "SYS_MODULE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaMODULE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", "SYS_MODULE");
            ILD_COMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_HR_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_HR_DEPT_ALL.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaHR_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaM_DEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_M_DEPT_MAPPING.SetLookupParamValue("W_MODULE_TYPE", W_MODULE_TYPE.EditValue);
            ILD_M_DEPT_MAPPING.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaM_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_M_DEPT_MAPPING.SetLookupParamValue("W_MODULE_TYPE", IGR_DEPT_MAPPING.GetCellValue("MODULE_TYPE"));
            ILD_M_DEPT_MAPPING.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_PERSON_VALUER_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", iDate.ISGetDate());
        }

        private void ILA_PERSON_VALUER_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", iDate.ISGetDate());
        }

        private void ILA_PERSON_LEADER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_PERSON.SetLookupParamValue("W_STD_DATE", iDate.ISGetDate());
        }
         
        #endregion

    }
}