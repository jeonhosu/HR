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

namespace EAPF0201
{
    public partial class EAPF0201 : Office2007Form
    {
        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public EAPF0201()
        {
            InitializeComponent();
        }

        public EAPF0201(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchFromDataAdapter()
        {
            IDA_CURRENCY.Fill();
        }

        #endregion;


        #region -- Default Value Setting --
        private void GRID_DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            IGR_CURRENCY.SetCellValue("EFFECTIVE_DATE_FR", idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE"));
            IGR_CURRENCY.SetCellValue("ENABLED_FLAG", "Y");
        }

        #endregion


        #region ----- Events -----

        private void EAPF0201_Load(object sender, EventArgs e)
        {
            IDA_CURRENCY.FillSchema();
        }

        private void isGridAdvEx1_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            int vColIndex_DateFrom = IDA_CURRENCY.OraSelectData.Columns.IndexOf("EFFECTIVE_DATE_FR"); //유효시작일
            int vColIndex_DateTo = IDA_CURRENCY.OraSelectData.Columns.IndexOf("EFFECTIVE_DATE_TO");   //유효종료일

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

            if (e.ColIndex == 1)
            {
               // isGridAdvEx1.GetCellValue("CURRENCY_CODE").
                string vText = e.NewValue.ToString();
                int vLength = vText.Length;

                if (vLength > 3)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10027", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                    e.Cancel = true;
                    return;
                }
            }
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
                    if (IDA_CURRENCY.IsFocused == true)
                    {
                        IDA_CURRENCY.AddOver();
                        GRID_DefaultValue();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_CURRENCY.IsFocused == true)
                    {
                        IDA_CURRENCY.AddUnder();
                        GRID_DefaultValue();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_CURRENCY.IsFocused == true)
                    {
                        IDA_CURRENCY.Update();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_CURRENCY.IsFocused == true)
                    {
                        IDA_CURRENCY.Cancel();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_CURRENCY.IsFocused == true)
                    {
                        IDA_CURRENCY.Delete();
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
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            } 
        }


    }
}