using System;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace EAPF0210
{
    public partial class EAPF0210 : Office2007Form
    {
        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public EAPF0210()
        {
            InitializeComponent();
        }

        public EAPF0210(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchFromDataAdapter()
        {
            isDataAdapter1.Fill();
        }

        #endregion;

        #region ----- Events -----

        private void EAPF0210_Load(object sender, EventArgs e)
        {
            isDataAdapter1.FillSchema();
            isDataAdapter2.FillSchema();
        }

        private void isAppInterfaceAdv1_AppMainButtonClick_1(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchFromDataAdapter();
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (isDataAdapter1.IsFocused == true)
                    {
                        isDataAdapter1.AddOver();
                    }
                    else if (isDataAdapter2.IsFocused == true)
                    {
                        isDataAdapter2.AddOver();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (isDataAdapter1.IsFocused == true)
                    {
                        isDataAdapter1.AddUnder();
                    }
                    else if (isDataAdapter2.IsFocused == true)
                    {
                        isDataAdapter2.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (isDataAdapter1.IsFocused == true)
                    {
                        isDataAdapter1.Update();
                    }
                    else if (isDataAdapter2.IsFocused == true)
                    {
                        isDataAdapter2.Update();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (isDataAdapter1.IsFocused == true)
                    {
                        isDataAdapter1.Cancel();
                    }
                    else if (isDataAdapter2.IsFocused == true)
                    {
                        isDataAdapter2.Cancel();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (isDataAdapter1.IsFocused == true)
                    {
                        isDataAdapter1.Delete();
                    }
                    else if (isDataAdapter2.IsFocused == true)
                    {
                        isDataAdapter2.Delete();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Print)
                {
                }
            }
        }
        #endregion;
    }
}