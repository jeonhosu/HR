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

namespace HRMF0111
{
    public partial class HRMF0111 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----
        public HRMF0111(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;

        #region ----- Private Methods ----
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
            if (iString.ISNull(W_STD_DATE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_DATE.Focus();
                return;
            }
            IDA_FOOD_MANAGER.Fill();
        }

        private void Insert_DB()
        {
            IGR_FOOD_MANAGER.SetCellValue("ENABLED_FLAG", "Y");
            IGR_FOOD_MANAGER.SetCellValue("EFFECTIVE_DATE_FR", iDate.ISMonth_1st(DateTime.Today));
            IGR_FOOD_MANAGER.Focus();
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----

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
                    if(IDA_FOOD_MANAGER.IsFocused)
                    {
                        IDA_FOOD_MANAGER.AddOver();
                        Insert_DB();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_FOOD_MANAGER.IsFocused)
                    {
                        IDA_FOOD_MANAGER.AddUnder();
                        Insert_DB();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_FOOD_MANAGER.Update(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_FOOD_MANAGER.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_FOOD_MANAGER.CurrentRow.RowState == DataRowState.Added)
                    {
                        IDA_FOOD_MANAGER.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ------ Form Event -----
        private void HRMF0111_Load(object sender, EventArgs e)
        {
            W_STD_DATE.EditValue = DateTime.Today;
            W_ENABLED_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;

            IDA_FOOD_MANAGER.FillSchema(); ;
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }
        #endregion

        #region ----- Adapter Event -----
        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["FOOD_DEVICE_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Device Name(장치명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["USER_ID"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=User Name(담당자명)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Effective Date From(유효 시작일)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

        #region ----- Lookup Event -----
        private void ILA_FOOD_DEVICE_W_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FOOD_DEVICE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_FOOD_DEVICE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FOOD_DEVICE.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }
        #endregion        
        
    }
}