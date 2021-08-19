using System;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using Syncfusion.Windows.Forms;
using InfoSummit.Win.ControlAdv;

namespace AppAssemblyEntry
{
    public partial class AppAssemblyEntry : Office2007Form
    {
        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public AppAssemblyEntry()
        {
            InitializeComponent();
        }

        public AppAssemblyEntry(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void EditAdvTextClear()
        {
            iedASSEMBLY_ID.EditValue = null;
        }

        #endregion;

        #region ----- Events -----

        private void EAPF0303_Load(object sender, EventArgs e)
        {
            isDataAdapter1.FillSchema();
        }

        private void btnOpenDlg_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string vFullName = openFileDialog1.FileName;
                int vCutStart = vFullName.LastIndexOf("\\") + 1;
                int vCutLength = vFullName.Length - vCutStart;
                string vFileName = vFullName.Substring(vCutStart, vCutLength);
                string vFilePath = vFullName.Substring(0, vCutStart);
                txtPath.EditValue = vFilePath;
            }
        }

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    isDataAdapter1.Fill();
                    EditAdvTextClear();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    isDataAdapter1.AddOver();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    isDataAdapter1.AddUnder();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    isDataAdapter1.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    isDataAdapter1.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    isDataAdapter1.Delete();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                }
            }
        }

        #endregion;
    }
}