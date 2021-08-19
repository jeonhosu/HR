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

namespace HRMF0113
{
    public partial class HRMF0113 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0113(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion

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
            IDA_DONATION_TYPE.Fill();
            IGR_DONATION_TYPE.Focus();
        }

        private void Init_Insert_Header()
        {
            IGR_DONATION_TYPE.SetCellValue("ENABLED_FLAG", "Y");
            IGR_DONATION_TYPE.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));

            IGR_DONATION_TYPE.Focus();
        }
         
        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Button Click -----

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
                    if (IDA_DONATION_TYPE.IsFocused)
                    {
                        IDA_DONATION_TYPE.AddOver();
                        Init_Insert_Header();
                    }
                    else if (IDA_DONATION_CARRIED_OVER.IsFocused)
                    {
                        IDA_DONATION_CARRIED_OVER.AddOver(); 
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_DONATION_TYPE.IsFocused)
                    {
                        IDA_DONATION_TYPE.AddUnder();
                        Init_Insert_Header();
                    }
                    else if (IDA_DONATION_CARRIED_OVER.IsFocused)
                    {
                        IDA_DONATION_CARRIED_OVER.AddUnder(); 
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                        IDA_DONATION_TYPE.Update();                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DONATION_TYPE.IsFocused)
                    {
                        IDA_DONATION_TYPE.Cancel();
                    }
                    else if (IDA_DONATION_CARRIED_OVER.IsFocused)
                    {
                        IDA_DONATION_CARRIED_OVER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DONATION_TYPE.IsFocused)
                    {
                        IDA_DONATION_TYPE.Delete();
                    }
                    else if (IDA_DONATION_CARRIED_OVER.IsFocused)
                    {
                        IDA_DONATION_CARRIED_OVER.Delete();
                    }
                }
            }
        }

        #endregion
        
        #region ----- Form Event -----

        private void HRMF0113_Load(object sender, EventArgs e)
        {
            IDA_DONATION_TYPE.FillSchema();
            IDA_DONATION_CARRIED_OVER.FillSchema();
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
        }

        private void HRMF0113_Shown(object sender, EventArgs e)
        {

        }

        #endregion 

        #region ----- Lookup Event ----- 
        
        private void ILA_YEAR_STR_FR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_YEAR_STR.SetLookupParamValue("W_END_YEAR", "3000");
        }

        private void ILA_YEAR_STR_TO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_YEAR_STR.SetLookupParamValue("W_START_YEAR", IGR_DONATION_CARRIED_OVER.GetCellValue("PERIOD_YYYY_FR"));
            ILD_YEAR_STR.SetLookupParamValue("W_END_YEAR", "3000");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_DONATION_TYPE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["DONATION_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10013"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //코드 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["DONATION_TYPE_NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10014"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_FR"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 시작일자 입력
                e.Cancel = true;
                return;
            }
            if (e.Row["EFFECTIVE_DATE_TO"] != DBNull.Value)
            {
                if (Convert.ToDateTime(e.Row["EFFECTIVE_DATE_FR"]) > Convert.ToDateTime(e.Row["EFFECTIVE_DATE_TO"]))
                {// 시작일자 ~ 종료일자
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 기간 검증 오류
                    e.Cancel = true;
                    return;
                }
            }
        }

        private void IDA_DONATION_TYPE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(데이터)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void IDA_DONATION_CARRIED_OVER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERIOD_YYYY_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("시작 {0}", isMessageAdapter1.ReturnText("FCM_10022")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["PERIOD_YYYY_TO"]) == string.Empty)
            {
                MessageBoxAdv.Show(string.Format("종료 {0}", isMessageAdapter1.ReturnText("FCM_10022")), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);  // 코드명 입력
                e.Cancel = true;
                return;
            }            
        }

        #endregion

        #region ----- KeyDown Event -----

        private void iedCODE_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        private void iedCODE_NAME_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                SEARCH_DB();
            }
        }

        #endregion


    }
}