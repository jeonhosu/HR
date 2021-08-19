using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0326
{
    public partial class HRMF0326_PRINT : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime(); 

        #endregion;

        #region ---- Return Value -----

        public object Get_Printer_Type
        {
            get
            {
                return V_PRINT_TYPE.EditValue;
            }

        }

        public object Get_Print_Preview
        {
            get
            {
                return CB_PRINT_PREVIEW.CheckBoxString;
            }

        }

        #endregion

        #region ----- Constructor -----

        public HRMF0326_PRINT(ISAppInterface pAppInterface)
        {
            InitializeComponent();
            isAppInterfaceAdv1.AppInterface = pAppInterface; 
        }
        #endregion;

        #region ----- Export File Name Methods ----

        private string SetExportFileName(string pExportFileName)
        {
            string vExportFileName = string.Empty;

            try
            {
                vExportFileName = pExportFileName;
                vExportFileName = vExportFileName.Replace("/", "_");
                vExportFileName = vExportFileName.Replace("\\", "_");
                vExportFileName = vExportFileName.Replace("*", "_");
                vExportFileName = vExportFileName.Replace("<", "_");
                vExportFileName = vExportFileName.Replace(">", "_");
                vExportFileName = vExportFileName.Replace("|", "_");
                vExportFileName = vExportFileName.Replace("?", "_");
                vExportFileName = vExportFileName.Replace(":", "_");
                vExportFileName = vExportFileName.Replace(" ", "_");
            }
            catch
            {
            }

            return vExportFileName;
        }


        #endregion;

        #region ----- Private Methods -----

        private void HRMF0326_PRINT_Load(object sender, EventArgs e)
        {

        }

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
         
        #region ----- Button Event ----

        private void isRadioButtonAdv_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;

            if (vRadio.Checked == true)
            {
                V_PRINT_TYPE.EditValue = vRadio.RadioCheckedString;
            }
        }

        // 발급 버튼 선택
        private void BTN_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISNull(V_PRINT_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10327"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            this.DialogResult = DialogResult.OK;
            this.Close(); 
        }

        // 취소 버튼 선택
        private void BTN_CLOSED_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion;

        #region ----- Lookup Event -----

        
        #endregion;

        #region ----- Adapter Event -----
         
        #endregion;


    }
}