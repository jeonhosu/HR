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

namespace HRMF0401
{
    public partial class HRMF0401 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        #region ----- Constructor -----
        public HRMF0401(Form pMainForm, ISAppInterface pAppInterface)
        {
            this.Visible = false;
            this.DoubleBuffered = true;

            InitializeComponent();

            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;

        #region ----- Property / Method -----

        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
        }

        private void SEARCH_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체 선택
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (STD_DATE_0.EditValue == null)
            {// 기준일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (TB_MAIN.SelectedTab.TabIndex == TP_NATIONAL_PENSION.TabIndex)
            {
                IDA_NATIONAL_PENSION.Fill();
                IGR_NATIONAL_PENSION.Focus();
            }
            else if (TB_MAIN.SelectedTab.TabIndex == TP_HEALTH_INSUR.TabIndex)
            {
                IDA_HEALTH_INSUR.Fill();
                IGR_HEALTH_INSUR.Focus();
            }
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

        #region ----- Excel Upload : Asset Master -----

        private void Select_Excel_File(string pINSUR_TYPE)
        {
            try
            {
                DirectoryInfo vOpenFolder = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                openFileDialog1.Title = "Select Open File";
                openFileDialog1.Filter = "Excel File(*.xls;*.xlsx)|*.xls;*.xlsx|All File(*.*)|*.*";
                openFileDialog1.DefaultExt = "xls";
                openFileDialog1.FileName = "*.xls;*.xlsx";

                string vFILE_PATH = string.Empty;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    vFILE_PATH = openFileDialog1.FileName;
                }
                else
                {
                    vFILE_PATH = string.Empty;
                }

                if (pINSUR_TYPE == "P")
                {
                    V_FILE_PATH_P.EditValue = vFILE_PATH;
                }
                else if (pINSUR_TYPE == "M")
                {
                    V_FILE_PATH_M.EditValue = vFILE_PATH;
                }
            }
            catch (Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                Application.DoEvents();
            }
        }

        //private void Excel_Upload_P()
        //{
        //    string vSTATUS = string.Empty;
        //    string vMESSAGE = string.Empty;
        //    bool vXL_Load_OK = false;

        //    if (iString.ISNull(V_FILE_PATH_P.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_FILE_PATH_P))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }
        //    Application.UseWaitCursor = true;
        //    this.Cursor = Cursors.WaitCursor;
        //    Application.DoEvents();

        //    string vOPenFileName = V_FILE_PATH_P.EditValue.ToString();
        //    XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);

        //    try
        //    {
        //        vXL_Upload.OpenFileName = vOPenFileName;
        //        vXL_Load_OK = vXL_Upload.OpenXL();
        //    }
        //    catch (Exception ex)
        //    {
        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);

        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();
        //        return;
        //    }


        //    ////기존자료 삭제.
        //    //vSTATUS = "F";
        //    //vMESSAGE = string.Empty;

        //    //IDC_DELETE_ASSET_MASTER_TEMP.ExecuteNonQuery();
        //    //vSTATUS = iString.ISNull(IDC_DELETE_ASSET_MASTER_TEMP.GetCommandParamValue("O_STATUS"));
        //    //vMESSAGE = iString.ISNull(IDC_DELETE_ASSET_MASTER_TEMP.GetCommandParamValue("O_MESSAGE"));
        //    //if (IDC_SET_TRANS_ASSET_MASTER.ExcuteError || vSTATUS == "F")
        //    //{
        //    //    Application.UseWaitCursor = false;
        //    //    this.Cursor = Cursors.Default;
        //    //    Application.DoEvents();

        //    //    if (vMESSAGE != string.Empty)
        //    //    {
        //    //        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    //    }
        //    //    return;

        //    //}

        //    // 업로드 아답터 fill //
        //    IDA_UPLOAD_PENSION.Cancel();
        //    IDA_UPLOAD_PENSION.Fill();   
        //    try
        //    {
        //        if (vXL_Load_OK == true)
        //        {
        //            vXL_Load_OK = vXL_Upload.LoadXL_P(IDA_UPLOAD_PENSION, 2);
        //            if (vXL_Load_OK == false)
        //            {
        //                IDA_UPLOAD_PENSION.Cancel();
        //            }
        //            else
        //            {
        //                IDA_UPLOAD_PENSION.Update();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        IDA_UPLOAD_PENSION.Cancel();

        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);

        //        vXL_Upload.DisposeXL();

        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();
        //        return;
        //    }
        //    vXL_Upload.DisposeXL();


        //    if (IDA_UPLOAD_PENSION.IsUpdateCompleted == true)
        //    {
        //        vSTATUS = "F";
        //        vMESSAGE = string.Empty;

        //        IDC_INTERFACE_INSUR_CHARGE.SetCommandParamValue("P_INSUR_TYPE", "P");
        //        IDC_INTERFACE_INSUR_CHARGE.ExecuteNonQuery();
        //        vSTATUS = iString.ISNull(IDC_INTERFACE_INSUR_CHARGE.GetCommandParamValue("O_STATUS"));
        //        vMESSAGE = iString.ISNull(IDC_INTERFACE_INSUR_CHARGE.GetCommandParamValue("O_MESSAGE"));

        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();
        //        if (IDC_INTERFACE_INSUR_CHARGE.ExcuteError || vSTATUS == "F")
        //        {
        //            if (vMESSAGE != string.Empty)
        //            {
        //                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            }
        //            return;
        //        }


        //        if (vSTATUS == "S")
        //        {
        //            if (vMESSAGE != string.Empty)
        //            {
        //                MessageBoxAdv.Show(vMESSAGE, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            }
        //        }
        //    }

        //    V_FILE_PATH_P.EditValue = string.Empty;
        //    Application.UseWaitCursor = false;
        //    this.Cursor = Cursors.Default;
        //    Application.DoEvents();
        //}

        //private void Excel_Upload_M()
        //{
        //    string vSTATUS = string.Empty;
        //    string vMESSAGE = string.Empty;
        //    bool vXL_Load_OK = false;

        //    if (iString.ISNull(V_FILE_PATH_M.EditValue) == string.Empty)
        //    {
        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", string.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(V_FILE_PATH_M))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }
        //    Application.UseWaitCursor = true;
        //    this.Cursor = Cursors.WaitCursor;
        //    Application.DoEvents();

        //    string vOPenFileName = V_FILE_PATH_M.EditValue.ToString();
        //    XL_Upload vXL_Upload = new XL_Upload(isAppInterfaceAdv1, isMessageAdapter1);

        //    try
        //    {
        //        vXL_Upload.OpenFileName = vOPenFileName;
        //        vXL_Load_OK = vXL_Upload.OpenXL();
        //    }
        //    catch (Exception ex)
        //    {
        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);

        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();
        //        return;
        //    }


        //    ////기존자료 삭제.
        //    //vSTATUS = "F";
        //    //vMESSAGE = string.Empty;

        //    //IDC_DELETE_ASSET_MASTER_TEMP.ExecuteNonQuery();
        //    //vSTATUS = iString.ISNull(IDC_DELETE_ASSET_MASTER_TEMP.GetCommandParamValue("O_STATUS"));
        //    //vMESSAGE = iString.ISNull(IDC_DELETE_ASSET_MASTER_TEMP.GetCommandParamValue("O_MESSAGE"));
        //    //if (IDC_SET_TRANS_ASSET_MASTER.ExcuteError || vSTATUS == "F")
        //    //{
        //    //    Application.UseWaitCursor = false;
        //    //    this.Cursor = Cursors.Default;
        //    //    Application.DoEvents();

        //    //    if (vMESSAGE != string.Empty)
        //    //    {
        //    //        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    //    }
        //    //    return;

        //    //}
            
        //    // 업로드 아답터 fill //
        //    IDA_UPLOAD_HEALTH.Cancel();
        //    IDA_UPLOAD_HEALTH.Fill();
        //    try
        //    {
        //        if (vXL_Load_OK == true)
        //        {
        //            vXL_Load_OK = vXL_Upload.LoadXL_M(IDA_UPLOAD_HEALTH, 2);
        //            if (vXL_Load_OK == false)
        //            {
        //                IDA_UPLOAD_HEALTH.Cancel();
        //            }
        //            else
        //            {
        //                IDA_UPLOAD_HEALTH.Update();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        IDA_UPLOAD_HEALTH.Cancel();

        //        isAppInterfaceAdv1.OnAppMessage(ex.Message);

        //        vXL_Upload.DisposeXL();

        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();
        //        return;
        //    }
        //    vXL_Upload.DisposeXL();


        //    if (IDA_UPLOAD_HEALTH.IsUpdateCompleted == true)
        //    {
        //        vSTATUS = "F";
        //        vMESSAGE = string.Empty;

        //        IDC_INTERFACE_INSUR_CHARGE.SetCommandParamValue("P_INSUR_TYPE", "M");
        //        IDC_INTERFACE_INSUR_CHARGE.ExecuteNonQuery();
        //        vSTATUS = iString.ISNull(IDC_INTERFACE_INSUR_CHARGE.GetCommandParamValue("O_STATUS"));
        //        vMESSAGE = iString.ISNull(IDC_INTERFACE_INSUR_CHARGE.GetCommandParamValue("O_MESSAGE"));

        //        Application.UseWaitCursor = false;
        //        this.Cursor = Cursors.Default;
        //        Application.DoEvents();
        //        if (IDC_INTERFACE_INSUR_CHARGE.ExcuteError || vSTATUS == "F")
        //        {
        //            if (vMESSAGE != string.Empty)
        //            {
        //                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            }
        //            return;
        //        }

        //        if (vSTATUS == "S")
        //        {
        //            if (vMESSAGE != string.Empty)
        //            {
        //                MessageBoxAdv.Show(vMESSAGE, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            }
        //        }
        //    }
        //    V_FILE_PATH_M.EditValue = string.Empty;
        //    Application.UseWaitCursor = false;
        //    this.Cursor = Cursors.Default;
        //    Application.DoEvents();
        //}

        #endregion

        #region ----- isAppInterfaceAdv1_AppMainButtonClick -----
        public void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_NATIONAL_PENSION.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_NATIONAL_PENSION.IsFocused)
                    {
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_NATIONAL_PENSION.Update();
                    IDA_HEALTH_INSUR.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_NATIONAL_PENSION.IsFocused)
                    {
                        IDA_NATIONAL_PENSION.Cancel();
                    }
                    else if (IDA_HEALTH_INSUR.IsFocused)
                    {
                        IDA_HEALTH_INSUR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
            }
        }
        #endregion

        #region ----- Form Event -----
        private void HRMF0401_Load(object sender, EventArgs e)
        {
            this.Visible = true;
            // FillSchema
            IDA_NATIONAL_PENSION.FillSchema();
            STD_DATE_0.EditValue = DateTime.Today;

            DefaultCorporation();                  // Corp Default Value Setting.
            
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]
        }

        private void BTN_FILE_SELECT_P_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File("P");
        }

        private void BTN_UPLOAD_EXEC_P_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0401_IMPORT vHRMF0401_IMPORT = new HRMF0401_IMPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface, CORP_ID_0.EditValue, "P");
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0401_IMPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0401_IMPORT.ShowDialog();
            vHRMF0401_IMPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                SEARCH_DB(); 
            }

            //Excel_Upload_P();   
        }

        private void BTN_FILE_SELECT_M_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Select_Excel_File("M");
        }

        private void BTN_UPLOAD_EXEC_M_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0401_IMPORT vHRMF0401_IMPORT = new HRMF0401_IMPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface, CORP_ID_0.EditValue, "M");
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0401_IMPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0401_IMPORT.ShowDialog();
            vHRMF0401_IMPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                SEARCH_DB();
            }
        }

        #endregion

        #region ----- Data Adapter Event ----

        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {// 사원.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["FLOOR_ID"] == DBNull.Value)
            {// FLOOR_ID
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10017"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CC_ID"] == DBNull.Value)
            {// cc
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10018"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

        }

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=해당 자료"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
        #endregion

        #region ----- Lookup Event -----
        private void ilaINSUR_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            //FLOOR
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "INSUR_TYPE");
            ildCOMMON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildCOMMON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaDEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            // DEPT
            ildDEPT.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildDEPT.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "N");
        }

        private void ilaPERSON_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            // PERSON
            ildPERSON.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ildPERSON.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);
        }
        #endregion


    }
}