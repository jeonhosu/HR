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

namespace HRMF0781
{
    public partial class HRMF0781_SET : Office2007Form
    {       

        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public object Set_Type
        {
            get
            {
                return SET_TYPE.EditValue;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public HRMF0781_SET(ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        
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

        private void HRMF0781_SET_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0781_SET_Shown(object sender, EventArgs e)
        {
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }

        private void isRadioButtonAdv1_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv iSTATUS = sender as ISRadioButtonAdv;
            SET_TYPE.EditValue = iSTATUS.RadioButtonString;
        }
    
        private void ibtAPPLY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void ibtCANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion              

        #region ----- Lookup Event -----

        #endregion

    }
}