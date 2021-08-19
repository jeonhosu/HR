using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace AppMenuEntry
{
    public partial class AppMenuEntry : Office2007Form
    {
        #region ----- Constructor -----

        public AppMenuEntry()
        {
            InitializeComponent();
        }

        public AppMenuEntry(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;        

        #region ----- Private Methods ----

        private void DataAdapter1_ExcuteQuery()
        {
            isDataAdapter1.SetSelectParamValue("W_MENU_NAME", isEditAdv3.EditValue);
            isDataAdapter1.Fill();
        }        

        #endregion;

        #region ----- Events -----

        private void APPF0050_Load(object sender, EventArgs e)
        {
            isLookupData1.SetLookupParamValue("W_MENU_NAME", "%");
            isDataAdapter1.FillSchema();
        }

        private void AppMenuEntry_Shown(object sender, EventArgs e)
        {
            isEditAdv3.Focus();
        }

        private void isDataAdapter1_ExcuteKeySearch(object pSender)
        {
            DataAdapter1_ExcuteQuery();
        }

        private void isGridAdvEx2_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            int vIndexCol1 = isGridAdvEx2.GetColumnToIndex("ENTRY_SEQ");
            int vIndexCol2 = isGridAdvEx2.GetColumnToIndex("ENTRY_PROMPT");
            if (e.ColIndex == vIndexCol1)
            {
                SendKeys.Send("+{TAB}"); 
                SendKeys.Send("{TAB}");
            }
        }

        private void isGridAdvEx2_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {

            int vIndexCol = isGridAdvEx2.GetColumnToIndex("ASSEMBLY_ID");
            if (e.ColIndex == vIndexCol)
            {
                isGridAdvEx2.SetCellValue("MENU_NAME", null);
                SendKeys.Send("{TAB}");
            }
        }

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    DataAdapter1_ExcuteQuery();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        isEditAdv3.LookupAdapter = null;
                        isDataAdapter1.AddOver();
                    }
                    else if (isDataAdapter2.IsFocused)
                    {
                        isDataAdapter2.AddOver();
                    }                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        isEditAdv3.LookupAdapter = null;
                        isDataAdapter1.AddUnder();
                    }
                    else if (isDataAdapter2.IsFocused)
                    {
                        isDataAdapter2.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        isDataAdapter1.SetInsertParamValue("P_CREATED_BY", 0);
                        isDataAdapter1.SetUpdateParamValue("P_LAST_UPDATED_BY", 0);
                        isDataAdapter1.Update();
                        isEditAdv3.LookupAdapter = isLookupAdapter1;
                    }
                    else if (isDataAdapter2.IsFocused)
                    {
                        isDataAdapter2.SetInsertParamValue("P_CREATED_BY", 0);
                        isDataAdapter2.SetUpdateParamValue("P_LAST_UPDATED_BY", 0);
                        isDataAdapter2.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        isEditAdv3.LookupAdapter = isLookupAdapter1;
                        isDataAdapter1.Cancel();                        
                    }
                    else if (isDataAdapter2.IsFocused)
                    {
                        isDataAdapter2.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (isDataAdapter1.IsFocused)
                    {
                        DialogResult ChoiceValue;

                        string vMessageString = string.Format("{0}", isMessageAdapter1.ReturnText("EAPP_10030"));
                        ChoiceValue = MessageBoxAdv.Show(vMessageString, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

                        if (ChoiceValue == DialogResult.Yes)
                        {
                            isDataAdapter1.Delete();
                        }
                    }
                    else if (isDataAdapter2.IsFocused)
                    {
                        isDataAdapter2.Delete();
                    }
                }
            }
        }              

        #endregion;
    }
}