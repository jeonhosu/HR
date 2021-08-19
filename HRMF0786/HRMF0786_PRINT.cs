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

namespace HRMF0786
{
    public partial class HRMF0786_PRINT : Office2007Form
    {       

        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public string Print_1_YN
        {
            //원천징수이행상황신고서
            get
            {
                return iString.ISNull(CB_PRINT_1.CheckBoxValue, "N");
            }
        }

        public string Print_2_YN
        {
            //소득세납부서-근로소득
            get
            {
                return iString.ISNull(CB_PRINT_2.CheckBoxValue, "N");
            }
        }

        #endregion;

        #region ----- Constructor -----

        public HRMF0786_PRINT(ISAppInterface pAppInterface)
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

        private void HRMF0786_PRINT_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0786_PRINT_Shown(object sender, EventArgs e)
        {
            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
        }
               
        private void ibtPRINTING_ButtonClick(object pSender, EventArgs pEventArgs)
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