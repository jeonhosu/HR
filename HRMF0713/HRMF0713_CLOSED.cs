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

namespace HRMF0713
{
    public partial class HRMF0713_CLOSED : Office2007Form
    {
        #region ----- Variables -----
        ISCommonUtil.ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISCommonUtil.ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();


        #endregion;

        #region ----- Constructor -----

        public HRMF0713_CLOSED()
        {
            InitializeComponent();
        }

        public HRMF0713_CLOSED(Form pMainForm, ISAppInterface pAppInterface
                                , object pSTD_YYYYMM
                                , object pClosed_Flag, object pClosed_Flag_Desc
                                , object pCorp_Desc, object pCorp_ID
                                , object pDept_Desc, object pDept_ID
                                , object pFloor_Desc, object pFloor_ID
                                , object pPerson_Name, object pPerson_Num, object pPerson_ID)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            W_STD_YYYYMM.EditValue = pSTD_YYYYMM;
            W_CORP_NAME.EditValue = pCorp_Desc;
            W_CORP_ID.EditValue = pCorp_ID;
            W_DEPT_NAME.EditValue = pDept_Desc;
            W_DEPT_ID.EditValue = pDept_ID;
            W_FLOOR_DESC.EditValue = pFloor_Desc;
            W_FLOOR_ID.EditValue = pFloor_ID;
            W_PERSON_NAME.EditValue = pPerson_Name;
            W_PERSON_NUM.EditValue = pPerson_Num;
            W_PERSON_ID.EditValue = pPerson_ID;
            W_CLOSED_FLAG_DESC.EditValue = pClosed_Flag_Desc;
            W_CLOSED_FLAG.EditValue = pClosed_Flag;

            if (iString.ISNull(pClosed_Flag) == "N")
            {
                this.Text = "Adjustment Closed";
            }
            else
            {
                this.Text = "Adjustment Cancel Closed";
            }
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchDB()
        {
            if (iString.ISNull(W_CORP_ID.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_CORP_NAME.Focus();
                return;
            }
            if (iString.ISNull(W_STD_YYYYMM.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10015"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                W_STD_YYYYMM.Focus();
                return;
            }
            
            IDA_YEAR_ADJUSTMENT_CLOSED.Fill();            
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchDB();
                    
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

        #region ----- Form Event ------
        
        private void HRMF0713_CLOSED_Load(object sender, EventArgs e)
        {
            
        }

        private void HRMF0713_CLOSED_Shown(object sender, EventArgs e)
        {
            W_CLOSED_FLAG_DESC.BringToFront();
        }

        private void BTN_SEARCH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            SearchDB();
        }

        private void BTN_SET_PROC_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            Application.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            string vSTATUS = "F";
            string vMESSAGE = string.Empty;

            int vIDX_SELECT_YN = IGR_YEAR_ADJUSTMENT_CLOSED.GetColumnToIndex("SELECT_YN");
            int vIDX_YEAR_YYYYMM = IGR_YEAR_ADJUSTMENT_CLOSED.GetColumnToIndex("YEAR_YYYYMM");
            int vIDX_PERSON_ID = IGR_YEAR_ADJUSTMENT_CLOSED.GetColumnToIndex("PERSON_ID");

            for (int vRow = 0; vRow < IGR_YEAR_ADJUSTMENT_CLOSED.RowCount; vRow++)
            {
                if (IGR_YEAR_ADJUSTMENT_CLOSED.GetCellValue(vRow, vIDX_SELECT_YN).ToString() == "Y")
                {
                    IGR_YEAR_ADJUSTMENT_CLOSED.CurrentCellMoveTo(vRow, vIDX_SELECT_YN);

                    IDC_SET_ADJUST_CLOSED.SetCommandParamValue("P_YEAR_YYYYMM", IGR_YEAR_ADJUSTMENT_CLOSED.GetCellValue(vRow, vIDX_YEAR_YYYYMM));
                    IDC_SET_ADJUST_CLOSED.SetCommandParamValue("P_PERSON_ID", IGR_YEAR_ADJUSTMENT_CLOSED.GetCellValue(vRow, vIDX_PERSON_ID));
                    IDC_SET_ADJUST_CLOSED.ExecuteNonQuery();
                    vSTATUS = iString.ISNull(IDC_SET_ADJUST_CLOSED.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iString.ISNull(IDC_SET_ADJUST_CLOSED.GetCommandParamValue("O_MESSAGE"));

                    if (IDC_SET_ADJUST_CLOSED.ExcuteError || vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        this.Cursor = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    //체크 해제 //
                    IGR_YEAR_ADJUSTMENT_CLOSED.SetCellValue(vRow, vIDX_SELECT_YN, "N");
                }
            }
            
            IGR_YEAR_ADJUSTMENT_CLOSED.LastConfirmChanges();
            IDA_YEAR_ADJUSTMENT_CLOSED.OraSelectData.AcceptChanges();
            IDA_YEAR_ADJUSTMENT_CLOSED.Refillable = true;

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMESSAGE);

            SearchDB();
        }

        private void BTN_CLOSED_FORM_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        private void IGR_YEAR_ADJUSTMENT_CLOSED_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (IGR_YEAR_ADJUSTMENT_CLOSED.RowIndex < 0)
            {
                return;
            }
            int vIDX_SELECT_YN = IGR_YEAR_ADJUSTMENT_CLOSED.GetColumnToIndex("SELECT_YN");
            if (e.ColIndex == vIDX_SELECT_YN)
            {
                IGR_YEAR_ADJUSTMENT_CLOSED.LastConfirmChanges();
                IDA_YEAR_ADJUSTMENT_CLOSED.OraSelectData.AcceptChanges();
                IDA_YEAR_ADJUSTMENT_CLOSED.Refillable = true;
            }  
        }

        private void CB_SELECT_YN_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            if (IGR_YEAR_ADJUSTMENT_CLOSED.RowCount < 1)
            {
                return;
            }
            
            int vIDX_SELECT_YN = IGR_YEAR_ADJUSTMENT_CLOSED.GetColumnToIndex("SELECT_YN");
            for (int vRow = 0; vRow < IGR_YEAR_ADJUSTMENT_CLOSED.RowCount; vRow++)
            {
                IGR_YEAR_ADJUSTMENT_CLOSED.SetCellValue(vRow, vIDX_SELECT_YN, CB_SELECT_YN.CheckBoxValue);
            }
            IGR_YEAR_ADJUSTMENT_CLOSED.LastConfirmChanges();
            IDA_YEAR_ADJUSTMENT_CLOSED.OraSelectData.AcceptChanges();
            IDA_YEAR_ADJUSTMENT_CLOSED.Refillable = true;
        }

        #endregion

        #region ----- Lookup Event -----

        private void ILA_W_FLOOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ILA_JOB_CATEGORY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "JOB_CATEGORY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP_0.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
        }

        private void ILA_W_YEAR_EMPLOYE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "YEAR_EMPLOYE_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}