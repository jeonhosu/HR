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

namespace EAPF0219
{
    public partial class EAPF0219 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public EAPF0219()
        {
            InitializeComponent();
        }

        public EAPF0219(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----


        #endregion;

        #region -- Default Value Setting ----

        //private void GRID_DefaultValue()
        //{
        //    idcLOCAL_DATE.ExecuteNonQuery();
        //    IGR_FTP.SetCellValue("EFFECTIVE_DATE_FR", idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE"));
        //    IGR_FTP.SetCellValue("ENABLED_FLAG", "Y");
        //}

        private void INSERT()
        {
            HOST_PORT.EditValue = 21;
            ENABLED_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;
            EFFECTIVE_DATE_FR.EditValue = iDate.ISMonth_1st(DateTime.Today);

            FTP_CODE.Focus();
        }

        #endregion

        #region ----- Territory Get Methods ----

        private int GetTerritory(ISUtil.Enum.TerritoryLanguage pTerritoryEnum)
        {
            int vTerritory = 0;

            switch (pTerritoryEnum)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    vTerritory = 1;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    vTerritory = 2;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    vTerritory = 3;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    vTerritory = 4;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    vTerritory = 5;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    vTerritory = 6;
                    break;
            }

            return vTerritory;
        }

        private object Get_Edit_Prompt(InfoSummit.Win.ControlAdv.ISEditAdv pEdit)
        {
            int mIDX = 0;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    mPrompt = pEdit.PromptTextElement[mIDX].Default;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL1_KR;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL2_CN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL3_VN;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL4_JP;
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    mPrompt = pEdit.PromptTextElement[mIDX].TL5_XAA;
                    break;
            }
            return mPrompt;
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick_1(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Search)
                {
                    IDA_FTP.Fill();
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_FTP.IsFocused == true)
                    {
                        IDA_FTP.AddOver();
                        INSERT();
                    } 
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_FTP.IsFocused == true)
                    {
                        IDA_FTP.AddUnder();
                        INSERT();
                    } 
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_FTP.Update(); 
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_FTP.IsFocused == true)
                    {
                        IDA_FTP.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_FTP.IsFocused == true)
                    {
                        if (IDA_FTP.CurrentRow.RowState == DataRowState.Added)
                        {
                            IDA_FTP.Delete();
                        }
                    } 
                } 
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Print)
                {
                }
            }
        }

        #endregion;

        #region ----- Form Event ----- 
        
        private void EAPF0219_Load(object sender, EventArgs e)
        {
            IDA_FTP.FillSchema();
        }
         
        #endregion

        #region ----- Adapter Event ----- 

        private void IDA_FTP_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(e.Row["FTP_CODE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FTP_CODE))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                FTP_CODE.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["FTP_DESC"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(FTP_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                FTP_DESC.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["HOST_IP"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(HOST_IP))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                HOST_IP.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["HOST_PORT"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(HOST_PORT))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                HOST_PORT.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["USER_NO"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(USER_NO))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                USER_NO.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["HOST_FOLDER"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(HOST_FOLDER))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                USER_NO.Focus();
                return;
            }
            if (iConv.ISNull(e.Row["EFFECTIVE_DATE_FR"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(EFFECTIVE_DATE_FR))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                EFFECTIVE_DATE_FR.Focus();
                return;
            }
        }

        #endregion
        
    }
}