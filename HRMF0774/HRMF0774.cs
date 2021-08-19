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

namespace HRMF0774
{
    public partial class HRMF0774 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0774()
        {
            InitializeComponent();
        }

        public HRMF0774(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        //诀眉
        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void Search_DB()
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iConv.ISNull(STD_DATE_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(STD_DATE_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_DATE_0.Focus();
                return;
            }
            IDA_RESIDENT_BUSINESS_ETC.Fill();
            IGR_RESIDENT_BUSINESS.Focus();
        }

        private void SetCommonParameter(object P_GROUP_CODE, object P_ENABLED_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", P_GROUP_CODE);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", P_ENABLED_YN);
        }

        #endregion;

        #region ----- 林家包府 ----

        private void Show_Address()
        {
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            DialogResult dlgRESULT;
            EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent, isAppInterfaceAdv1.AppInterface, ZIP_CODE.EditValue, ADDR1.EditValue);
            dlgRESULT = vEAPF0299.ShowDialog();

            if (dlgRESULT == DialogResult.OK)
            {
                ZIP_CODE.EditValue = vEAPF0299.Get_Zip_Code;
                ADDR1.EditValue = vEAPF0299.Get_Address;
            }
            vEAPF0299.Dispose();
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        //private void Show_Address_OPERATING_UNIT(int pIDX_Row, int pIDX_ZIP_CODE, int pIDX_ADDRESS)
        //{
        //    Application.UseWaitCursor = true;
        //    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //    Application.DoEvents();

        //    DialogResult dlgRESULT;
        //    igrHRM_OPERATING_UNIT_G.LastConfirmChanges();
        //    EAPF0299.EAPF0299 vEAPF0299 = new EAPF0299.EAPF0299(this.MdiParent
        //                                            , isAppInterfaceAdv1.AppInterface
        //                                            , igrHRM_OPERATING_UNIT_G.GetCellValue(pIDX_Row, pIDX_ZIP_CODE)
        //                                            , igrHRM_OPERATING_UNIT_G.GetCellValue(pIDX_Row, pIDX_ADDRESS));
        //    dlgRESULT = vEAPF0299.ShowDialog();

        //    if (dlgRESULT == DialogResult.OK)
        //    {
        //        igrHRM_OPERATING_UNIT_G.SetCellValue(pIDX_Row, pIDX_ZIP_CODE, vEAPF0299.Get_Zip_Code);
        //        igrHRM_OPERATING_UNIT_G.SetCellValue(pIDX_Row, pIDX_ADDRESS, vEAPF0299.Get_Address);
        //    }
        //    vEAPF0299.Dispose();
        //    this.Cursor = System.Windows.Forms.Cursors.Default;
        //    Application.UseWaitCursor = false;
        //    Application.DoEvents();
        //}

        #endregion

        #region ----- 林刮锅龋 眉农 -----

        private object REPRE_NUM_Check(object pRepre_num)
        {
            object Check_YN = "N";
            if (iConv.ISNull(pRepre_num) == string.Empty)
            {
                return Check_YN;
            }
                        
            idcREPRE_NUM_CHECK.SetCommandParamValue("P_REPRE_NUM", pRepre_num);
            idcREPRE_NUM_CHECK.ExecuteNonQuery();
            Check_YN = idcREPRE_NUM_CHECK.GetCommandParamValue("O_RETURN_VALUE");
            return Check_YN;
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
            try
            {                
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
            }
            catch
            {
            }
            return mPrompt;
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    Search_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_RESIDENT_BUSINESS_ETC.IsFocused)
                    {
                        IDA_RESIDENT_BUSINESS_ETC.AddOver();
                        NAME.Focus();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_RESIDENT_BUSINESS_ETC.IsFocused)
                    {
                        IDA_RESIDENT_BUSINESS_ETC.AddUnder();
                        NAME.Focus();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {                    
                    IDA_RESIDENT_BUSINESS_ETC.Update();
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_RESIDENT_BUSINESS_ETC.IsFocused)
                    { 
                        IDA_RESIDENT_BUSINESS_ETC.Cancel();
                    } 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    //if (IDA_RESIDENT_BUSINESS.IsFocused)
                    //{
                    //    if (IGR_RESIDENT_BSN_FAMILY.RowCount > 0)
                    //    {
                    //        IDA_RESIDENT_BSN_FAMILY.MoveFirst(IGR_RESIDENT_BSN_FAMILY.Name);
                    //        for (int C = 0; C < IGR_RESIDENT_BSN_FAMILY.RowCount; C++)
                    //        {
                    //            IDA_RESIDENT_BSN_FAMILY.Delete();
                    //            IDA_RESIDENT_BSN_FAMILY.MoveNext(IGR_RESIDENT_BSN_FAMILY.Name);
                    //        }
                    //    }
                    //    IDA_RESIDENT_BUSINESS.Delete();
                    //} 
                }
            }
        }

        #endregion;

        #region ----- Form event -----

        private void HRMF0774_Load(object sender, EventArgs e)
        {
            IDA_RESIDENT_BUSINESS_ETC.FillSchema(); 
        }

        private void HRMF0774_Shown(object sender, EventArgs e)
        {
            DefaultCorporation();
            STD_DATE_0.EditValue = DateTime.Today;
        }

        private void ZIP_CODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address();
            }
        }

        private void ADDRESS1_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Show_Address();
            }
        }

        #endregion

        #region ----- Lookup event -----

        private void ilaNATIONALITY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("NATIONALITY_TYPE", "Y");
        }

        private void ilaFLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("FLOOR", "Y");
        }

        private void ilaBUSINESS_CODE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BUSINESS_CODE", "Y");
        }

        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("BANK", "Y");
        }

        private void ilaRELATION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("RELATION", "Y");
        }
         
        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_RESIDENT_BUSINESS_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(CORP_NAME_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["NAME"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["REPRE_NUM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(REPRE_NUM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull(e.Row["NATIONALITY_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(NATIONALITY_TYPE_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            } 
        }
         
        #endregion




    }
}