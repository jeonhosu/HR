using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;       //호환되지 않은DLL을 사용할 때.

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace HRMF0781
{
    public partial class HRMF0781_SUB : Office2007Form
    {  
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object mWITHHOLDING_DOC_ID;
 
        #endregion;

        #region ----- Constructor -----

        public HRMF0781_SUB()
        {
            InitializeComponent();
        }

        public HRMF0781_SUB(Form pMainForm, ISAppInterface pAppInterface, object pWITHHOLDING_DOC_ID)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            mWITHHOLDING_DOC_ID = pWITHHOLDING_DOC_ID;
        }

        #endregion;

        #region ----- Private Methods ----
          
        private void Search_DB()
        {
            if (iConv.ISNull(mWITHHOLDING_DOC_ID) == string.Empty)
            {
                return;
            }

            IDA_WITHHOLDING_DOC_SUB_01.OraSelectData.AcceptChanges();
            IDA_WITHHOLDING_DOC_SUB_01.Refillable = true;
            IGR_WITHHOLDING_DOC_SUB_01.ResetDraw = true;

            IDA_WITHHOLDING_DOC_SUB_01.SetSelectParamValue("P_WITHHOLDING_DOC_ID", mWITHHOLDING_DOC_ID);
            IDA_WITHHOLDING_DOC_SUB_01.Fill();
            Sync_WITHHOLDING_DOC_SUB_01();

            IDA_WITHHOLDING_DOC_SUB_02.OraSelectData.AcceptChanges();
            IDA_WITHHOLDING_DOC_SUB_02.Refillable = true;
            IGR_WITHHOLDING_DOC_SUB_02.ResetDraw = true;

            IDA_WITHHOLDING_DOC_SUB_02.SetSelectParamValue("P_WITHHOLDING_DOC_ID", mWITHHOLDING_DOC_ID);
            IDA_WITHHOLDING_DOC_SUB_02.Fill();
            Sync_WITHHOLDING_DOC_SUB_02();

            IDA_WITHHOLDING_DOC_SUB_03.OraSelectData.AcceptChanges();
            IDA_WITHHOLDING_DOC_SUB_03.Refillable = true;
            IGR_WITHHOLDING_DOC_SUB_03.ResetDraw = true;

            IDA_WITHHOLDING_DOC_SUB_03.SetSelectParamValue("P_WITHHOLDING_DOC_ID", mWITHHOLDING_DOC_ID);
            IDA_WITHHOLDING_DOC_SUB_03.Fill();
            Sync_WITHHOLDING_DOC_SUB_03();
        }

        private void Sync_WITHHOLDING_DOC_SUB_01()
        {
            string vPAYMENT_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue("PAYMENT_FLAG"));
            string vSUMMARY_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue("SUMMARY_FLAG"));
            string vINCOME_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue("INCOME_FLAG"));
            string vSP_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue("SP_FLAG"));
            string vADD_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue("ADD_FLAG"));
            string vREFUND_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue("REFUND_FLAG"));

            //금액//
            int vIDX_PERSON_CNT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("PERSON_CNT");
            int vIDX_PAYMENT_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("PAYMENT_AMT");
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("REFUND_TAX_AMT");

            int Insertable = 0;
            int Updatable = 0;

            //인원.
            if(vSUMMARY_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_PERSON_CNT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_PERSON_CNT].Updatable = Updatable;

            //소득 총지급액.
            if (vSUMMARY_FLAG == "Y" || vPAYMENT_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_PAYMENT_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_PAYMENT_AMT].Updatable = Updatable;

            //소득세.
            if (vSUMMARY_FLAG == "Y" || vINCOME_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_INCOME_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_INCOME_TAX_AMT].Updatable = Updatable;

            //농특세.
            if (vSUMMARY_FLAG == "Y" || vSP_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_SP_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_SP_TAX_AMT].Updatable = Updatable;

            //가산세.
            if (vSUMMARY_FLAG == "Y" || vADD_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_ADD_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_ADD_TAX_AMT].Updatable = Updatable;

            //환급세액.
            if (vSUMMARY_FLAG == "Y" || vREFUND_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_REFUND_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_01.GridAdvExColElement[vIDX_REFUND_TAX_AMT].Updatable = Updatable;

            IGR_WITHHOLDING_DOC_SUB_01.Refresh();
        }


        private void Sync_WITHHOLDING_DOC_SUB_02()
        {
            string vPAYMENT_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue("PAYMENT_FLAG"));
            string vSUMMARY_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue("SUMMARY_FLAG"));
            string vINCOME_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue("INCOME_FLAG"));
            string vSP_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue("SP_FLAG"));
            string vADD_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue("ADD_FLAG"));
            string vREFUND_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue("REFUND_FLAG"));

            //금액//
            int vIDX_PERSON_CNT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("PERSON_CNT");
            int vIDX_PAYMENT_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("PAYMENT_AMT");
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("REFUND_TAX_AMT");

            int Insertable = 0;
            int Updatable = 0;

            //인원.
            if (vSUMMARY_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_PERSON_CNT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_PERSON_CNT].Updatable = Updatable;

            //소득 총지급액.
            if (vSUMMARY_FLAG == "Y" || vPAYMENT_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_PAYMENT_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_PAYMENT_AMT].Updatable = Updatable;

            //소득세.
            if (vSUMMARY_FLAG == "Y" || vINCOME_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_INCOME_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_INCOME_TAX_AMT].Updatable = Updatable;

            //농특세.
            if (vSUMMARY_FLAG == "Y" || vSP_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_SP_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_SP_TAX_AMT].Updatable = Updatable;

            //가산세.
            if (vSUMMARY_FLAG == "Y" || vADD_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_ADD_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_ADD_TAX_AMT].Updatable = Updatable;

            //환급세액.
            if (vSUMMARY_FLAG == "Y" || vREFUND_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_REFUND_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_02.GridAdvExColElement[vIDX_REFUND_TAX_AMT].Updatable = Updatable;

            IGR_WITHHOLDING_DOC_SUB_02.Refresh();
        }

        private void Sync_WITHHOLDING_DOC_SUB_03()
        {
            string vPAYMENT_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue("PAYMENT_FLAG"));
            string vSUMMARY_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue("SUMMARY_FLAG"));
            string vINCOME_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue("INCOME_FLAG"));
            string vSP_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue("SP_FLAG"));
            string vADD_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue("ADD_FLAG"));
            string vREFUND_FLAG = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue("REFUND_FLAG"));

            //금액//
            int vIDX_PERSON_CNT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("PERSON_CNT");
            int vIDX_PAYMENT_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("PAYMENT_AMT");
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("REFUND_TAX_AMT");

            int Insertable = 0;
            int Updatable = 0;

            //인원.
            if (vSUMMARY_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_PERSON_CNT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_PERSON_CNT].Updatable = Updatable;

            //소득 총지급액.
            if (vSUMMARY_FLAG == "Y" || vPAYMENT_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_PAYMENT_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_PAYMENT_AMT].Updatable = Updatable;

            //소득세.
            if (vSUMMARY_FLAG == "Y" || vINCOME_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_INCOME_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_INCOME_TAX_AMT].Updatable = Updatable;

            //농특세.
            if (vSUMMARY_FLAG == "Y" || vSP_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_SP_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_SP_TAX_AMT].Updatable = Updatable;

            //가산세.
            if (vSUMMARY_FLAG == "Y" || vADD_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_ADD_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_ADD_TAX_AMT].Updatable = Updatable;

            //환급세액.
            if (vSUMMARY_FLAG == "Y" || vREFUND_FLAG == "Y")
            {
                Insertable = 0;
                Updatable = 0;
            }
            else
            {
                Insertable = 1;
                Updatable = 1;
            }
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_REFUND_TAX_AMT].Insertable = Insertable;
            IGR_WITHHOLDING_DOC_SUB_03.GridAdvExColElement[vIDX_REFUND_TAX_AMT].Updatable = Updatable;

            IGR_WITHHOLDING_DOC_SUB_03.Refresh();
        }

        #endregion;

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
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_WITHHOLDING_DOC_SUB_01.Update(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_WITHHOLDING_DOC_SUB_01.IsFocused)
                    {
                        IDA_WITHHOLDING_DOC_SUB_01.Cancel();
                    }
                    else if (IDA_WITHHOLDING_DOC_SUB_02.IsFocused)
                    {
                        IDA_WITHHOLDING_DOC_SUB_02.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                     
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {

                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0781_SUB_Load(object sender, EventArgs e)
        {
             
        }

        private void HRMF0781_SUB_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized; // 화면 최대화//
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IDA_WITHHOLDING_DOC_SUB_01.FillSchema();
            IDA_WITHHOLDING_DOC_SUB_02.FillSchema();
            IDA_WITHHOLDING_DOC_SUB_03.FillSchema();

            Search_DB();
        }

        private void IGR_WITHHOLDING_LIST_CellDoubleClick(object pSender)
        {
             
        }

        private void BTN_INQUIRY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Search_DB();
        }
        
        private void BTN_SAVE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_WITHHOLDING_DOC_SUB_01.Update();
            IDA_WITHHOLDING_DOC_SUB_02.Update();
            IDA_WITHHOLDING_DOC_SUB_03.Update();
        }

        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if(IDA_WITHHOLDING_DOC_SUB_01.IsFocused)
            {
                IDA_WITHHOLDING_DOC_SUB_01.Cancel();
            }
            else if (IDA_WITHHOLDING_DOC_SUB_02.IsFocused)
            {
                IDA_WITHHOLDING_DOC_SUB_02.Cancel();
            }
            else if (IDA_WITHHOLDING_DOC_SUB_03.IsFocused)
            {
                IDA_WITHHOLDING_DOC_SUB_03.Cancel();
            }
        }

        private void BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void IGR_WITHHOLDING_DOC_SUB_01_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if(IGR_WITHHOLDING_DOC_SUB_01.RowCount < 1)
            {
                return;
            }

            int vIDX_PERSON_COUNT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("PERSON_CNT");
            int vIDX_PAYMENT_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("PAYMENT_AMT");
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("REFUND_TAX_AMT");
            int vIDX_FIX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("FIX_INCOME_TAX_AMT");
            int vIDX_FIX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("FIX_SP_TAX_AMT");

            string vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue("INCOME_SUB_GROUP_CODE"));

            decimal vFIX_INCOME_TAX_AMT = 0;
            decimal vFIX_SP_TAX_AMT = 0;

            if (e.ColIndex == vIDX_PERSON_COUNT)
            {
                SUM_PERSON_CNT_01(vIDX_PERSON_COUNT, e.NewValue);
            }
            else if(e.ColIndex == vIDX_PAYMENT_AMT)
            {
                SUM_PAYMENT_AMT_01(vIDX_PAYMENT_AMT, e.NewValue);
            }
            else if (e.ColIndex == vIDX_INCOME_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(e.NewValue, 0) 
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_SP_TAX_AMT), 0);  
                if(vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if(vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
               
                SUM_INCOME_TAX_AMT_01(e.RowIndex, e.NewValue); 
                //SUM_FIX_INCOME_TAX_AMT_01(e.RowIndex, vFIX_INCOME_TAX_AMT);
                //SUM_FIX_SP_TAX_AMT_01(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_SP_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(e.NewValue, 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_SP_TAX_AMT_01(e.RowIndex, e.NewValue);
            //    SUM_FIX_INCOME_TAX_AMT_01(e.RowIndex, vFIX_INCOME_TAX_AMT);
            //    SUM_FIX_SP_TAX_AMT_01(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_ADD_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(e.NewValue, 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_SP_TAX_AMT), 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_ADD_TAX_AMT_01(e.RowIndex, e.NewValue);
                //SUM_FIX_INCOME_TAX_AMT_01(e.RowIndex, vFIX_INCOME_TAX_AMT);
                //SUM_FIX_SP_TAX_AMT_01(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_REFUND_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(e.NewValue, 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_SP_TAX_AMT), 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_REFUND_TAX_AMT_01(e.RowIndex, e.NewValue); 
            } 
        }


        private void IGR_WITHHOLDING_DOC_SUB_02_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (IGR_WITHHOLDING_DOC_SUB_02.RowCount < 1)
            {
                return;
            }

            int vIDX_PERSON_COUNT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("PERSON_CNT");
            int vIDX_PAYMENT_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("PAYMENT_AMT");
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("REFUND_TAX_AMT");
            int vIDX_FIX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("FIX_INCOME_TAX_AMT");
            int vIDX_FIX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("FIX_SP_TAX_AMT");

            string vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue("INCOME_SUB_GROUP_CODE"));

            decimal vFIX_INCOME_TAX_AMT = 0;
            decimal vFIX_SP_TAX_AMT = 0;

            if (e.ColIndex == vIDX_PERSON_COUNT)
            {
                SUM_PERSON_CNT_02(vIDX_PERSON_COUNT, e.NewValue);
            }
            else if (e.ColIndex == vIDX_PAYMENT_AMT)
            {
                SUM_PAYMENT_AMT_02(vIDX_PAYMENT_AMT, e.NewValue);
            }
            else if (e.ColIndex == vIDX_INCOME_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(e.NewValue, 0)
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_SP_TAX_AMT), 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);

                SUM_INCOME_TAX_AMT_02(e.RowIndex, e.NewValue);
                //SUM_FIX_INCOME_TAX_AMT_02(e.RowIndex, vFIX_INCOME_TAX_AMT);
                //SUM_FIX_SP_TAX_AMT_02(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_SP_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(e.NewValue, 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_SP_TAX_AMT_02(e.RowIndex, e.NewValue);
                //    SUM_FIX_INCOME_TAX_AMT_02(e.RowIndex, vFIX_INCOME_TAX_AMT);
                //    SUM_FIX_SP_TAX_AMT_02(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_ADD_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(e.NewValue, 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_SP_TAX_AMT), 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_ADD_TAX_AMT_02(e.RowIndex, e.NewValue);
                //SUM_FIX_INCOME_TAX_AMT_02(e.RowIndex, vFIX_INCOME_TAX_AMT);
                //SUM_FIX_SP_TAX_AMT_02(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_REFUND_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(e.NewValue, 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(vIDX_SP_TAX_AMT), 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_REFUND_TAX_AMT_02(e.RowIndex, e.NewValue);
            }
        }

        private void IGR_WITHHOLDING_DOC_SUB_03_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (IGR_WITHHOLDING_DOC_SUB_03.RowCount < 1)
            {
                return;
            }

            int vIDX_PERSON_COUNT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("PERSON_CNT");
            int vIDX_PAYMENT_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("PAYMENT_AMT");
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("REFUND_TAX_AMT");
            int vIDX_FIX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("FIX_INCOME_TAX_AMT");
            int vIDX_FIX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("FIX_SP_TAX_AMT");

            string vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue("INCOME_SUB_GROUP_CODE"));

            decimal vFIX_INCOME_TAX_AMT = 0;
            decimal vFIX_SP_TAX_AMT = 0;

            if (e.ColIndex == vIDX_PERSON_COUNT)
            {
                SUM_PERSON_CNT_03(vIDX_PERSON_COUNT, e.NewValue);
            }
            else if (e.ColIndex == vIDX_PAYMENT_AMT)
            {
                SUM_PAYMENT_AMT_03(vIDX_PAYMENT_AMT, e.NewValue);
            }
            else if (e.ColIndex == vIDX_INCOME_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(e.NewValue, 0)
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_SP_TAX_AMT), 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);

                SUM_INCOME_TAX_AMT_03(e.RowIndex, e.NewValue);
                //SUM_FIX_INCOME_TAX_AMT_03(e.RowIndex, vFIX_INCOME_TAX_AMT);
                //SUM_FIX_SP_TAX_AMT_03(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_SP_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(e.NewValue, 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_SP_TAX_AMT_03(e.RowIndex, e.NewValue);
                //    SUM_FIX_INCOME_TAX_AMT_03(e.RowIndex, vFIX_INCOME_TAX_AMT);
                //    SUM_FIX_SP_TAX_AMT_03(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_ADD_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(e.NewValue, 0)
                                    - iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_REFUND_TAX_AMT), 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_SP_TAX_AMT), 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_ADD_TAX_AMT_03(e.RowIndex, e.NewValue);
                //SUM_FIX_INCOME_TAX_AMT_03(e.RowIndex, vFIX_INCOME_TAX_AMT);
                //SUM_FIX_SP_TAX_AMT_03(e.RowIndex, vFIX_SP_TAX_AMT);
            }
            else if (e.ColIndex == vIDX_REFUND_TAX_AMT)
            {
                vFIX_INCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_INCOME_TAX_AMT), 0)
                                    + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_ADD_TAX_AMT), 0)
                                    - iConv.ISDecimaltoZero(e.NewValue, 0);
                vFIX_SP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(vIDX_SP_TAX_AMT), 0);
                if (vFIX_INCOME_TAX_AMT < 0)
                {
                    vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                    if (vFIX_SP_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = 0;
                    }
                    vFIX_INCOME_TAX_AMT = 0;
                }
                IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(e.RowIndex, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(e.RowIndex, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                SUM_REFUND_TAX_AMT_03(e.RowIndex, e.NewValue);
            }
        }

        #endregion

        #region ----- Lookup Event -----

        #endregion

        #region ----- Adapter event ------

        private void IDA_WITHHOLDING_DOC_SUB_01_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Sync_WITHHOLDING_DOC_SUB_01();
        }

        private void IDA_WITHHOLDING_DOC_SUB_02_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Sync_WITHHOLDING_DOC_SUB_02();
        }

        private void IDA_WITHHOLDING_DOC_SUB_03_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Sync_WITHHOLDING_DOC_SUB_03();
        }

        #endregion

        #region ----- Sum or Total  :: 01 -----

        private void SUM_PERSON_CNT_01(int pRowIndex, object pValue)
        {
            decimal vC30_PERSON_COUNT = 0;
            decimal vC50_PERSON_COUNT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC30_RowIndex = 0;
            int vC50_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("PERSON_CNT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C30")
                {
                    if (r == pRowIndex)
                    {
                        vC30_PERSON_COUNT = vC30_PERSON_COUNT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC30_PERSON_COUNT = vC30_PERSON_COUNT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }
                if (vINCOME_SUB_GROUP_CODE == "50")
                {
                    if (r == pRowIndex)
                    {
                        vC50_PERSON_COUNT = vC50_PERSON_COUNT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC50_PERSON_COUNT = vC50_PERSON_COUNT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                } 

                if(iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C30")
                {
                    vC30_RowIndex = r;
                }
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C50")
                {
                    vC50_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC30_RowIndex, vIDX, vC30_PERSON_COUNT);
            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC50_RowIndex, vIDX, vC50_PERSON_COUNT); 
        }

        private void SUM_PAYMENT_AMT_01(int pRowIndex, object pValue)
        {
            decimal vC30_PAYMENT_AMT = 0;
            decimal vC50_PAYMENT_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC30_RowIndex = 0;
            int vC50_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("PAYMENT_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C30")
                {
                    if (r == pRowIndex)
                    {
                        vC30_PAYMENT_AMT = vC30_PAYMENT_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC30_PAYMENT_AMT = vC30_PAYMENT_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }
                if (vINCOME_SUB_GROUP_CODE == "50")
                {
                    if (r == pRowIndex)
                    {
                        vC50_PAYMENT_AMT = vC50_PAYMENT_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC50_PAYMENT_AMT = vC50_PAYMENT_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C30")
                {
                    vC30_RowIndex = r;
                }
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C50")
                {
                    vC50_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC30_RowIndex, vIDX, vC30_PAYMENT_AMT);
            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC50_RowIndex, vIDX, vC50_PAYMENT_AMT); 
        }

        private void SUM_INCOME_TAX_AMT_01(int pRowIndex, object pValue)
        {
            decimal vC30_INCOME_TAX_AMT = 0;
            decimal vC50_INCOME_TAX_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC30_RowIndex = 0;
            int vC50_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C30")
                {
                    if (r == pRowIndex)
                    {
                        vC30_INCOME_TAX_AMT = vC30_INCOME_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC30_INCOME_TAX_AMT = vC30_INCOME_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }
                if (vINCOME_SUB_GROUP_CODE == "50")
                {
                    if (r == pRowIndex)
                    {
                        vC50_INCOME_TAX_AMT = vC50_INCOME_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC50_INCOME_TAX_AMT = vC50_INCOME_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C30")
                {
                    vC30_RowIndex = r;
                }
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C50")
                {
                    vC50_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC30_RowIndex, vIDX, vC30_INCOME_TAX_AMT);
            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC50_RowIndex, vIDX, vC50_INCOME_TAX_AMT);

            SUM_FIX_TAX_AMT_01();
        }

        private void SUM_SP_TAX_AMT_01(int pRowIndex, object pValue)
        {
            decimal vC30_SP_TAX_AMT = 0;
            decimal vC50_SP_TAX_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC30_RowIndex = 0;
            int vC50_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("SP_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C30")
                {
                    if (r == pRowIndex)
                    {
                        vC30_SP_TAX_AMT = vC30_SP_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC30_SP_TAX_AMT = vC30_SP_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }
                if (vINCOME_SUB_GROUP_CODE == "50")
                {
                    if (r == pRowIndex)
                    {
                        vC50_SP_TAX_AMT = vC50_SP_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC50_SP_TAX_AMT = vC50_SP_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C30")
                {
                    vC30_RowIndex = r;
                }
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C50")
                {
                    vC50_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC30_RowIndex, vIDX, vC30_SP_TAX_AMT);
            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC50_RowIndex, vIDX, vC50_SP_TAX_AMT);

            SUM_FIX_TAX_AMT_01();
        }

        private void SUM_ADD_TAX_AMT_01(int pRowIndex, object pValue)
        {
            decimal vC30_TAX_AMT = 0;
            decimal vC50_TAX_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC30_RowIndex = 0;
            int vC50_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("ADD_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C30")
                {
                    if (r == pRowIndex)
                    {
                        vC30_TAX_AMT = vC30_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC30_TAX_AMT = vC30_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }
                if (vINCOME_SUB_GROUP_CODE == "50")
                {
                    if (r == pRowIndex)
                    {
                        vC50_TAX_AMT = vC50_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC50_TAX_AMT = vC50_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C30")
                {
                    vC30_RowIndex = r;
                }
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C50")
                {
                    vC50_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC30_RowIndex, vIDX, vC30_TAX_AMT);
            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC50_RowIndex, vIDX, vC50_TAX_AMT);

            SUM_FIX_TAX_AMT_01();
        }

        private void SUM_REFUND_TAX_AMT_01(int pRowIndex, object pValue)
        {
            decimal vC30_REFUND_TAX_AMT = 0;
            decimal vC50_REFUND_TAX_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC30_RowIndex = 0;
            int vC50_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("REFUND_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C30")
                {
                    if (r == pRowIndex)
                    {
                        vC30_REFUND_TAX_AMT = vC30_REFUND_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC30_REFUND_TAX_AMT = vC30_REFUND_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }
                if (vINCOME_SUB_GROUP_CODE == "50")
                {
                    if (r == pRowIndex)
                    {
                        vC50_REFUND_TAX_AMT = vC50_REFUND_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC50_REFUND_TAX_AMT = vC50_REFUND_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C30")
                {
                    vC30_RowIndex = r;
                }
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C50")
                {
                    vC50_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC30_RowIndex, vIDX, vC30_REFUND_TAX_AMT);
            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC50_RowIndex, vIDX, vC50_REFUND_TAX_AMT);

            SUM_FIX_TAX_AMT_01();
        }
            
        private bool CHECK_C30_PAY_TAX_AMT()
        {
            //납부세액 검증 
            //decimal vTOTOAL_PAY_TAX_AMT = 0;

            ////납부세액-소득세등(가산세 포함) 
            //if (iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue) > 0)
            //{
            //    vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue);
            //}
            //if (iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue) > 0)
            //{
            //    vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
            //                            iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue);
            //}
            ////납부세액-농특세            
            //if (iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue) > 0)
            //{//납부할 세액이 있는경우
            //    vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
            //                            iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue);
            //}

            ////납부세액보다 당월조정 환급세액이 많음
            //if ((iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue) +
            //    iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) +
            //    iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            //{
            //    MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등 + (11)농어촌특별세)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}
            return true;
        }

        private void SUM_FIX_TAX_AMT_01()
        {
            decimal vINCOME_TAX_AMT = 0;
            decimal vSP_TAX_AMT = 0;
            decimal vADD_TAX_AMT = 0;
            decimal vREFUND_TAX_AMT = 0;
            decimal vFIX_INCOME_TAX_AMT = 0;
            decimal vFIX_SP_TAX_AMT = 0;

            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");  
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("REFUND_TAX_AMT");
            int vIDX_FIX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("FIX_INCOME_TAX_AMT");
            int vIDX_FIX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("FIX_SP_TAX_AMT");
            IGR_WITHHOLDING_DOC_SUB_01.LastConfirmChanges();
            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE)) == "SUM")
                {
                    vINCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_INCOME_TAX_AMT));
                    vSP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_SP_TAX_AMT));
                    vADD_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_ADD_TAX_AMT));
                    vREFUND_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_REFUND_TAX_AMT));

                    vFIX_INCOME_TAX_AMT = vINCOME_TAX_AMT + vADD_TAX_AMT - vREFUND_TAX_AMT;
                    vFIX_SP_TAX_AMT = vSP_TAX_AMT;
                    if (vFIX_INCOME_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                        if (vFIX_SP_TAX_AMT < 0)
                        {
                            vFIX_SP_TAX_AMT = 0;
                        }
                        vFIX_INCOME_TAX_AMT = 0;
                    }
                    IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(r, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                    IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(r, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                } 
            }

            
        }
         
        private void SUM_FIX_INCOME_TAX_AMT_01(int pRowIndex, object pValue)
        {
            decimal vC30_AMT = 0;
            decimal vC50_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC30_RowIndex = 0;
            int vC50_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("FIX_INCOME_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C30")
                {
                    if (r == pRowIndex)
                    {
                        vC30_AMT = vC30_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC30_AMT = vC30_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }
                if (vINCOME_SUB_GROUP_CODE == "50")
                {
                    if (r == pRowIndex)
                    {
                        vC50_AMT = vC50_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC50_AMT = vC50_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C30")
                {
                    vC30_RowIndex = r;
                }
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C50")
                {
                    vC50_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC30_RowIndex, vIDX, vC30_AMT);
            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC50_RowIndex, vIDX, vC50_AMT);
        }

        private void SUM_FIX_SP_TAX_AMT_01(int pRowIndex, object pValue)
        {
            decimal vC30_AMT = 0;
            decimal vC50_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC30_RowIndex = 0;
            int vC50_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_01.GetColumnToIndex("FIX_SP_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_01.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C30")
                {
                    if (r == pRowIndex)
                    {
                        vC30_AMT = vC30_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC30_AMT = vC30_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }
                if (vINCOME_SUB_GROUP_CODE == "50")
                {
                    if (r == pRowIndex)
                    {
                        vC50_AMT = vC50_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC50_AMT = vC50_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C30")
                {
                    vC30_RowIndex = r;
                }
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_01.GetCellValue(r, vIDX_CODE)) == "C50")
                {
                    vC50_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC30_RowIndex, vIDX, vC30_AMT);
            IGR_WITHHOLDING_DOC_SUB_01.SetCellValue(vC50_RowIndex, vIDX, vC50_AMT);
        }

        #endregion


        #region ----- Sum or Total  :: 02 -----

        private void SUM_PERSON_CNT_02(int pRowIndex, object pValue)
        {
            decimal vC70_PERSON_COUNT = 0; 
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC70_RowIndex = 0; 
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("PERSON_CNT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C70")
                {
                    if (r == pRowIndex)
                    {
                        vC70_PERSON_COUNT = vC70_PERSON_COUNT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC70_PERSON_COUNT = vC70_PERSON_COUNT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_CODE)) == "C70")
                {
                    vC70_RowIndex = r;
                } 
            }
            IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(vC70_RowIndex, vIDX, vC70_PERSON_COUNT); 
        }

        private void SUM_PAYMENT_AMT_02(int pRowIndex, object pValue)
        {
            decimal vC70_PAYMENT_AMT = 0; 
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC70_RowIndex = 0; 
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("PAYMENT_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C70")
                {
                    if (r == pRowIndex)
                    {
                        vC70_PAYMENT_AMT = vC70_PAYMENT_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC70_PAYMENT_AMT = vC70_PAYMENT_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX));
                    }
                } 

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_CODE)) == "C70")
                {
                    vC70_RowIndex = r;
                } 
            }

            IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(vC70_RowIndex, vIDX, vC70_PAYMENT_AMT); 
        }

        private void SUM_INCOME_TAX_AMT_02(int pRowIndex, object pValue)
        {
            decimal vC70_INCOME_TAX_AMT = 0; 
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC70_RowIndex = 0; 
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C70")
                {
                    if (r == pRowIndex)
                    {
                        vC70_INCOME_TAX_AMT = vC70_INCOME_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC70_INCOME_TAX_AMT = vC70_INCOME_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX));
                    }
                } 

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_CODE)) == "C70")
                {
                    vC70_RowIndex = r;
                } 
            }

            IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(vC70_RowIndex, vIDX, vC70_INCOME_TAX_AMT); 

            SUM_FIX_TAX_AMT_02();
        }

        private void SUM_SP_TAX_AMT_02(int pRowIndex, object pValue)
        {
            decimal vC70_SP_TAX_AMT = 0; 
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC70_RowIndex = 0; 
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("SP_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C70")
                {
                    if (r == pRowIndex)
                    {
                        vC70_SP_TAX_AMT = vC70_SP_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC70_SP_TAX_AMT = vC70_SP_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX));
                    }
                } 

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_CODE)) == "C70")
                {
                    vC70_RowIndex = r;
                } 
            }

            IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(vC70_RowIndex, vIDX, vC70_SP_TAX_AMT); 

            SUM_FIX_TAX_AMT_02();
        }

        private void SUM_ADD_TAX_AMT_02(int pRowIndex, object pValue)
        {
            decimal vC70_TAX_AMT = 0; 
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC70_RowIndex = 0; 
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("ADD_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C70")
                {
                    if (r == pRowIndex)
                    {
                        vC70_TAX_AMT = vC70_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC70_TAX_AMT = vC70_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX));
                    }
                } 

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_CODE)) == "C70")
                {
                    vC70_RowIndex = r;
                } 
            }

            IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(vC70_RowIndex, vIDX, vC70_TAX_AMT); 

            SUM_FIX_TAX_AMT_02();
        }

        private void SUM_REFUND_TAX_AMT_02(int pRowIndex, object pValue)
        {
            decimal vC70_REFUND_TAX_AMT = 0; 
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC70_RowIndex = 0; 
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("REFUND_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C70")
                {
                    if (r == pRowIndex)
                    {
                        vC70_REFUND_TAX_AMT = vC70_REFUND_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC70_REFUND_TAX_AMT = vC70_REFUND_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX));
                    }
                } 

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_CODE)) == "C70")
                {
                    vC70_RowIndex = r;
                } 
            }

            IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(vC70_RowIndex, vIDX, vC70_REFUND_TAX_AMT); 

            SUM_FIX_TAX_AMT_02();
        }

        private bool CHECK_C70_PAY_TAX_AMT()
        {
            //납부세액 검증 
            //decimal vTOTOAL_PAY_TAX_AMT = 0;

            ////납부세액-소득세등(가산세 포함) 
            //if (iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue) > 0)
            //{
            //    vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue);
            //}
            //if (iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue) > 0)
            //{
            //    vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
            //                            iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue);
            //}
            ////납부세액-농특세            
            //if (iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue) > 0)
            //{//납부할 세액이 있는경우
            //    vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
            //                            iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue);
            //}

            ////납부세액보다 당월조정 환급세액이 많음
            //if ((iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue) +
            //    iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) +
            //    iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            //{
            //    MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등 + (11)농어촌특별세)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}
            return true;
        }

        private void SUM_FIX_TAX_AMT_02()
        {
            decimal vINCOME_TAX_AMT = 0;
            decimal vSP_TAX_AMT = 0;
            decimal vADD_TAX_AMT = 0;
            decimal vREFUND_TAX_AMT = 0;
            decimal vFIX_INCOME_TAX_AMT = 0;
            decimal vFIX_SP_TAX_AMT = 0;

            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("REFUND_TAX_AMT");
            int vIDX_FIX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("FIX_INCOME_TAX_AMT");
            int vIDX_FIX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("FIX_SP_TAX_AMT");
            IGR_WITHHOLDING_DOC_SUB_02.LastConfirmChanges();
            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE)) == "SUM")
                {
                    vINCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_INCOME_TAX_AMT));
                    vSP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_SP_TAX_AMT));
                    vADD_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_ADD_TAX_AMT));
                    vREFUND_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_REFUND_TAX_AMT));

                    vFIX_INCOME_TAX_AMT = vINCOME_TAX_AMT + vADD_TAX_AMT - vREFUND_TAX_AMT;
                    vFIX_SP_TAX_AMT = vSP_TAX_AMT;
                    if (vFIX_INCOME_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                        if (vFIX_SP_TAX_AMT < 0)
                        {
                            vFIX_SP_TAX_AMT = 0;
                        }
                        vFIX_INCOME_TAX_AMT = 0;
                    }
                    IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(r, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                    IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(r, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                }
            }


        }

        private void SUM_FIX_INCOME_TAX_AMT_02(int pRowIndex, object pValue)
        {
            decimal vC70_AMT = 0; 
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC70_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("FIX_INCOME_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C70")
                {
                    if (r == pRowIndex)
                    {
                        vC70_AMT = vC70_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC70_AMT = vC70_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX));
                    }
                } 

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_CODE)) == "C70")
                {
                    vC70_RowIndex = r;
                } 
            }

            IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(vC70_RowIndex, vIDX, vC70_AMT); 
        }

        private void SUM_FIX_SP_TAX_AMT_02(int pRowIndex, object pValue)
        {
            decimal vC70_AMT = 0; 
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC70_RowIndex = 0; 
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_02.GetColumnToIndex("FIX_SP_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_02.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C70")
                {
                    if (r == pRowIndex)
                    {
                        vC70_AMT = vC70_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC70_AMT = vC70_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX));
                    }
                } 

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_02.GetCellValue(r, vIDX_CODE)) == "C70")
                {
                    vC70_RowIndex = r;
                } 
            }

            IGR_WITHHOLDING_DOC_SUB_02.SetCellValue(vC70_RowIndex, vIDX, vC70_AMT); 
        }

        #endregion

        #region ----- Sum or Total  :: 03 -----

        private void SUM_PERSON_CNT_03(int pRowIndex, object pValue)
        {
            decimal vC90_PERSON_COUNT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC90_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("PERSON_CNT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C90")
                {
                    if (r == pRowIndex)
                    {
                        vC90_PERSON_COUNT = vC90_PERSON_COUNT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC90_PERSON_COUNT = vC90_PERSON_COUNT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_CODE)) == "C90")
                {
                    vC90_RowIndex = r;
                }
            }
            IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(vC90_RowIndex, vIDX, vC90_PERSON_COUNT);
        }

        private void SUM_PAYMENT_AMT_03(int pRowIndex, object pValue)
        {
            decimal vC90_PAYMENT_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC90_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("PAYMENT_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C90")
                {
                    if (r == pRowIndex)
                    {
                        vC90_PAYMENT_AMT = vC90_PAYMENT_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC90_PAYMENT_AMT = vC90_PAYMENT_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_CODE)) == "C90")
                {
                    vC90_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(vC90_RowIndex, vIDX, vC90_PAYMENT_AMT);
        }

        private void SUM_INCOME_TAX_AMT_03(int pRowIndex, object pValue)
        {
            decimal vC90_INCOME_TAX_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC90_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C90")
                {
                    if (r == pRowIndex)
                    {
                        vC90_INCOME_TAX_AMT = vC90_INCOME_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC90_INCOME_TAX_AMT = vC90_INCOME_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_CODE)) == "C90")
                {
                    vC90_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(vC90_RowIndex, vIDX, vC90_INCOME_TAX_AMT);

            SUM_FIX_TAX_AMT_03();
        }

        private void SUM_SP_TAX_AMT_03(int pRowIndex, object pValue)
        {
            decimal vC90_SP_TAX_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC90_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("SP_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C90")
                {
                    if (r == pRowIndex)
                    {
                        vC90_SP_TAX_AMT = vC90_SP_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC90_SP_TAX_AMT = vC90_SP_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_CODE)) == "C90")
                {
                    vC90_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(vC90_RowIndex, vIDX, vC90_SP_TAX_AMT);

            SUM_FIX_TAX_AMT_03();
        }

        private void SUM_ADD_TAX_AMT_03(int pRowIndex, object pValue)
        {
            decimal vC90_TAX_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC90_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("ADD_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C90")
                {
                    if (r == pRowIndex)
                    {
                        vC90_TAX_AMT = vC90_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC90_TAX_AMT = vC90_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_CODE)) == "C90")
                {
                    vC90_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(vC90_RowIndex, vIDX, vC90_TAX_AMT);

            SUM_FIX_TAX_AMT_03();
        }

        private void SUM_REFUND_TAX_AMT_03(int pRowIndex, object pValue)
        {
            decimal vC90_REFUND_TAX_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC90_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("REFUND_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C90")
                {
                    if (r == pRowIndex)
                    {
                        vC90_REFUND_TAX_AMT = vC90_REFUND_TAX_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC90_REFUND_TAX_AMT = vC90_REFUND_TAX_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_CODE)) == "C90")
                {
                    vC90_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(vC90_RowIndex, vIDX, vC90_REFUND_TAX_AMT);

            SUM_FIX_TAX_AMT_03();
        }

        private bool CHECK_C90_PAY_TAX_AMT()
        {
            //납부세액 검증 
            //decimal vTOTOAL_PAY_TAX_AMT = 0;

            ////납부세액-소득세등(가산세 포함) 
            //if (iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue) > 0)
            //{
            //    vTOTOAL_PAY_TAX_AMT = iConv.ISDecimaltoZero(A30_INCOME_TAX_AMT.EditValue);
            //}
            //if (iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue) > 0)
            //{
            //    vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
            //                            iConv.ISDecimaltoZero(A30_ADD_TAX_AMT.EditValue);
            //}
            ////납부세액-농특세            
            //if (iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue) > 0)
            //{//납부할 세액이 있는경우
            //    vTOTOAL_PAY_TAX_AMT = vTOTOAL_PAY_TAX_AMT +
            //                            iConv.ISDecimaltoZero(A30_SP_TAX_AMT.EditValue);
            //}

            ////납부세액보다 당월조정 환급세액이 많음
            //if ((iConv.ISDecimaltoZero(A30_THIS_REFUND_TAX_AMT.EditValue) +
            //    iConv.ISDecimaltoZero(A30_PAY_INCOME_TAX_AMT.EditValue) +
            //    iConv.ISDecimaltoZero(A30_PAY_SP_TAX_AMT.EditValue)) != vTOTOAL_PAY_TAX_AMT)
            //{
            //    MessageBoxAdv.Show("징수세액합계와 ((9)당월조정환급세액 + 납부세액((10)소득세등 + (11)농어촌특별세)합계 금액이 다릅니다. 확인하세요", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return false;
            //}
            return true;
        }

        private void SUM_FIX_TAX_AMT_03()
        {
            decimal vINCOME_TAX_AMT = 0;
            decimal vSP_TAX_AMT = 0;
            decimal vADD_TAX_AMT = 0;
            decimal vREFUND_TAX_AMT = 0;
            decimal vFIX_INCOME_TAX_AMT = 0;
            decimal vFIX_SP_TAX_AMT = 0;

            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_TAX_AMT");
            int vIDX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("SP_TAX_AMT");
            int vIDX_ADD_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("ADD_TAX_AMT");
            int vIDX_REFUND_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("REFUND_TAX_AMT");
            int vIDX_FIX_INCOME_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("FIX_INCOME_TAX_AMT");
            int vIDX_FIX_SP_TAX_AMT = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("FIX_SP_TAX_AMT");
            IGR_WITHHOLDING_DOC_SUB_03.LastConfirmChanges();
            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE)) == "SUM")
                {
                    vINCOME_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_INCOME_TAX_AMT));
                    vSP_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_SP_TAX_AMT));
                    vADD_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_ADD_TAX_AMT));
                    vREFUND_TAX_AMT = iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_REFUND_TAX_AMT));

                    vFIX_INCOME_TAX_AMT = vINCOME_TAX_AMT + vADD_TAX_AMT - vREFUND_TAX_AMT;
                    vFIX_SP_TAX_AMT = vSP_TAX_AMT;
                    if (vFIX_INCOME_TAX_AMT < 0)
                    {
                        vFIX_SP_TAX_AMT = vFIX_SP_TAX_AMT + vFIX_INCOME_TAX_AMT;
                        if (vFIX_SP_TAX_AMT < 0)
                        {
                            vFIX_SP_TAX_AMT = 0;
                        }
                        vFIX_INCOME_TAX_AMT = 0;
                    }
                    IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(r, vIDX_FIX_INCOME_TAX_AMT, vFIX_INCOME_TAX_AMT);
                    IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(r, vIDX_FIX_SP_TAX_AMT, vFIX_SP_TAX_AMT);
                }
            } 
        }

        private void SUM_FIX_INCOME_TAX_AMT_03(int pRowIndex, object pValue)
        {
            decimal vC90_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC90_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("FIX_INCOME_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C90")
                {
                    if (r == pRowIndex)
                    {
                        vC90_AMT = vC90_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC90_AMT = vC90_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_CODE)) == "C90")
                {
                    vC90_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(vC90_RowIndex, vIDX, vC90_AMT);
        }

        private void SUM_FIX_SP_TAX_AMT_03(int pRowIndex, object pValue)
        {
            decimal vC90_AMT = 0;
            string vINCOME_SUB_GROUP_CODE = string.Empty;

            int vC90_RowIndex = 0;
            int vIDX_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_CODE");
            int vIDX_GROUP_CODE = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("INCOME_SUB_GROUP_CODE");
            int vIDX = IGR_WITHHOLDING_DOC_SUB_03.GetColumnToIndex("FIX_SP_TAX_AMT");

            for (int r = 0; r < IGR_WITHHOLDING_DOC_SUB_03.RowCount; r++)
            {
                vINCOME_SUB_GROUP_CODE = iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_GROUP_CODE));
                if (vINCOME_SUB_GROUP_CODE == "C90")
                {
                    if (r == pRowIndex)
                    {
                        vC90_AMT = vC90_AMT + iConv.ISDecimaltoZero(pValue);
                    }
                    else
                    {
                        vC90_AMT = vC90_AMT + iConv.ISDecimaltoZero(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX));
                    }
                }

                if (iConv.ISNull(IGR_WITHHOLDING_DOC_SUB_03.GetCellValue(r, vIDX_CODE)) == "C90")
                {
                    vC90_RowIndex = r;
                }
            }

            IGR_WITHHOLDING_DOC_SUB_03.SetCellValue(vC90_RowIndex, vIDX, vC90_AMT);
        }

        #endregion

    }
}