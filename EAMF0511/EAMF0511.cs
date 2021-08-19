using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace EAMF0511
{
    public partial class EAMF0511 : Office2007Form
    {
        #region ----- Variables -----
    

        #endregion;

        #region ----- Constructor -----

        public EAMF0511()
        {
            InitializeComponent();
        }

        public EAMF0511(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            IDA_SULBI_MASTER.Fill();
            ISG_SULBI_MASTER.Focus();

        }


        #endregion;

        #region ----- Events -----

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
                    IDA_SULBI_MASTER.AddOver();

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    IDA_SULBI_MASTER.AddUnder();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {

                    IDA_SULBI_MASTER.Update();

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_SULBI_MASTER.IsFocused)
                    {
                        IDA_SULBI_MASTER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_SULBI_MASTER.IsFocused)
                    {
                        IDA_SULBI_MASTER.Delete();
                    }
                }
            }
        }
    
        #endregion;

    }
}