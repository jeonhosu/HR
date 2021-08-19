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

namespace HRMF0504
{
    public partial class HRMF0504_DETAIL : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;
        
        #region ----- Constructor -----

        public HRMF0504_DETAIL(Form pMainForm, ISAppInterface pAppInterface, object pPERIOD_NAME, object pPERSON_NAME, object pPERSON_ID)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_PERIOD_NAME.EditValue = pPERIOD_NAME;
            W_PERSON_NAME.EditValue = pPERSON_NAME;
            W_PERSON_ID.EditValue = pPERSON_ID;
        }

        #endregion;

        #region ----- Private Methods -----

        private void Search_DB()
        {
            IDA_GENERAL_HOURLY_AMT.Fill();
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----

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

        private void HRMF0504_DETAIL_Shown(object sender, EventArgs e)
        {
            Search_DB();
        }

        private void BTN_GEN_HOURLY_DETAIL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion

    }
}