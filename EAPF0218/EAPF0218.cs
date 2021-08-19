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
using System.IO;

namespace EAPF0218
{
    public partial class EAPF0218 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public EAPF0218()
        {
            InitializeComponent();
        }

        public EAPF0218(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----


        #endregion;

        #region -- Default Value Setting --
        private void GRID_DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            ISG_EMAIL_ENTRY.SetCellValue("EFFECTIVE_DATE_FR", idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE"));
            ISG_EMAIL_ENTRY.SetCellValue("ENABLED_FLAG", "Y");
        }

        #endregion

        #region ----- Events -----

        private void EAPF0218_Load(object sender, EventArgs e)
        {
            IDA_EMAIL_TYPE.FillSchema();
            IDA_EMAIL_ENTRY.FillSchema();
        }
        private void isGridAdvEx2_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            int vColIndex_DateFrom = IDA_EMAIL_ENTRY.OraSelectData.Columns.IndexOf("EFFECTIVE_DATE_FR"); //유효시작일
            int vColIndex_DateTo = IDA_EMAIL_ENTRY.OraSelectData.Columns.IndexOf("EFFECTIVE_DATE_TO");   //유효종료일

            if (e.ColIndex == vColIndex_DateTo)
            {
                string vTextDate = e.NewValue.ToString();
                bool isNull = string.IsNullOrEmpty(vTextDate);
                if (e.NewValue != null && isNull == false)
                {
                    ISGridAdvEx vGridAdvEx = pSender as ISGridAdvEx;
                    DateTime vDateFrom = (DateTime)vGridAdvEx.GetCellValue(vColIndex_DateFrom);
                    DateTime vDateTo = (DateTime)e.NewValue;

                    if (vDateFrom > vDateTo)
                    {
                        e.Cancel = true;

                        string vMessageString = string.Format("[{0}]~[{1}]\n{2}", vDateFrom, vDateTo, isMessageAdapter1.ReturnText("FCM_10012"));
                        MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void isAppInterfaceAdv1_AppMainButtonClick_1(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Search)
                {
                    IDA_EMAIL_TYPE.Fill();
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_EMAIL_TYPE.IsFocused == true)
                    {
                        IDA_EMAIL_TYPE.AddOver();
                    }
                    else if (IDA_EMAIL_ENTRY.IsFocused == true)
                    {
                        IDA_EMAIL_ENTRY.AddOver();
                        GRID_DefaultValue();
                    }
                    else if (IDA_EMAIL_MANAGER.IsFocused == true)
                    {
                        IDA_EMAIL_MANAGER.AddOver();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_EMAIL_TYPE.IsFocused == true)
                    {
                        IDA_EMAIL_TYPE.AddUnder();
                    }
                    else if (IDA_EMAIL_ENTRY.IsFocused == true)
                    {
                        IDA_EMAIL_ENTRY.AddUnder();
                        GRID_DefaultValue();
                    }
                    else if (IDA_EMAIL_MANAGER.IsFocused == true)
                    {
                        IDA_EMAIL_MANAGER.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_EMAIL_TYPE.IsFocused == true)
                    {
                        IDA_EMAIL_TYPE.Update();
                    }
                    else if (IDA_EMAIL_ENTRY.IsFocused == true)
                    {
                        IDA_EMAIL_ENTRY.Update();
                    }
                    else if (IDA_EMAIL_MANAGER.IsFocused == true)
                    {
                        IDA_EMAIL_MANAGER.Update();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_EMAIL_TYPE.IsFocused == true)
                    {
                        IDA_EMAIL_TYPE.Cancel();
                    }
                    else if (IDA_EMAIL_ENTRY.IsFocused == true)
                    {
                        IDA_EMAIL_ENTRY.Cancel();
                    }
                    else if (IDA_EMAIL_MANAGER.IsFocused == true)
                    {
                        IDA_EMAIL_MANAGER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_EMAIL_TYPE.IsFocused == true)
                    {
                        IDA_EMAIL_TYPE.Delete();
                    }
                    else if (IDA_EMAIL_ENTRY.IsFocused == true)
                    {
                        IDA_EMAIL_ENTRY.Delete();
                    }
                    else if (IDA_EMAIL_MANAGER.IsFocused == true)
                    {
                        IDA_EMAIL_MANAGER.Delete();
                    }
                } 
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Print)
                {
                }
            }
        }
        #endregion;

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
        {
            //if (e.Row.RowState != DataRowState.Added)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
            //    e.Cancel = true;
            //    return;
            //} 
        }

        private void isDataAdapter2_PreDelete(ISPreDeleteEventArgs e)
        {
            //if (e.Row.RowState != DataRowState.Added)
            //{
            //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
            //    e.Cancel = true;
            //    return;
            //} 
        }

        private void isLookupAdapter4_SelectedRowData(object pSender)
        {
            ISG_EMAIL_ENTRY.SetCellValue("NAME", "");
            ISG_EMAIL_ENTRY.SetCellValue("PERSON_NUM", "");
            ISG_EMAIL_ENTRY.SetCellValue("DISPLAY_NAME", "");
        }

        private void IDA_EMAIL_ENTRY_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(ISG_EMAIL_ENTRY.GetCellValue("EMAIL_PERSON_NAME")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10126"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }

            if (iConv.ISNull(ISG_EMAIL_ENTRY.GetCellValue("EMAIL_ADDRESS")) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10127"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                return;
            }
        }

 

    }
}