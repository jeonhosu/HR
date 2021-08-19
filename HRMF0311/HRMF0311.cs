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

namespace HRMF0311
{
    public partial class HRMF0311 : Office2007Form
    {
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF0311(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)   //파견직관리
            {
                G_CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
        }

        #endregion;

        #region ----- Corp Type -----

        private void V_RB_ALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            G_CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

        #endregion

        #region ----- Private Methods -----

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
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "N");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "N");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            CORP_NAME_0.BringToFront();
            //CORP TYPE :: 전체이면 그룹박스 표시, 
            if(iConv.ISNull(G_CORP_TYPE.EditValue, "1") == "1")
            {
                V_CORP_GROUP.Visible = false; //.Show();
                V_RB_OWNER.CheckedState = ISUtil.Enum.CheckedState.Checked;
                G_CORP_TYPE.EditValue = V_RB_OWNER.RadioCheckedString;
            }
            else
            {
                V_CORP_GROUP.Visible = true; //.Show();
                if (iConv.ISNull(G_CORP_TYPE.EditValue) == "ALL")
                {
                    V_RB_ALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G_CORP_TYPE.EditValue = V_RB_ALL.RadioCheckedString;
                }
                else
                {
                    V_RB_ETC.CheckedState = ISUtil.Enum.CheckedState.Checked;
                    G_CORP_TYPE.EditValue = V_RB_ETC.RadioCheckedString;
                }
            }
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_0.Focus();
                return;
            }
            this.UseWaitCursor = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            IDA_DAY_INTERFACE_TRANS.Fill();
            IGR_DAY_INTERFACE.Focus();

            this.Cursor = Cursors.Default;
            this.UseWaitCursor = false;
        }

        private void isSearch_WorkCalendar(Object pPerson_ID, Object pWork_Date)
        {
            ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
            if (iConvert.ISNull(pWork_Date) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = WORK_DATE_0.EditValue;

            idaDAY_HISTORY.Fill();
            idaDUTY_PERIOD.Fill();
        }

        private void isSearch_Day_History(int pAdd_Day)
        {
            ISFunction.ISConvert iConvert = new ISFunction.ISConvert();
            if (iConvert.ISNull(WORK_DATE_8.EditValue) == string.Empty)
            {
                return;
            }
            WORK_DATE_8.EditValue = Convert.ToDateTime(WORK_DATE_8.EditValue).AddDays(pAdd_Day);
            idaDAY_HISTORY.Fill();
        }


        private void Set_Update_Approve(object pApproved_Flag)
        {
            if (IGR_DAY_INTERFACE.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();
             
            int vIDX_PERSON_ID = IGR_DAY_INTERFACE.GetColumnToIndex("PERSON_ID");
            int vIDX_WORK_DATE = IGR_DAY_INTERFACE.GetColumnToIndex("WORK_DATE");
            int vIDX_CORP_ID = IGR_DAY_INTERFACE.GetColumnToIndex("CORP_ID");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_DAY_INTERFACE.RowCount; i++)
            {
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_PERSON_ID", IGR_DAY_INTERFACE.GetCellValue(i, vIDX_PERSON_ID));
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_WORK_DATE", IGR_DAY_INTERFACE.GetCellValue(i, vIDX_WORK_DATE));
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_CORP_ID", IGR_DAY_INTERFACE.GetCellValue(i, vIDX_CORP_ID));
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_STATUS", "C");
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_CHECK_YN", "Y");
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_FLAG", pApproved_Flag); 
                IDC_SET_UPDATE_APPROVE.ExecuteNonQuery();  
                vSTATUS = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_MESSAGE"));
                if (IDC_SET_UPDATE_APPROVE.ExcuteError || vSTATUS == "F")
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    Application.DoEvents();
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                } 
            }
             
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        private void Set_Update_Review(object pApproved_Flag)
        {
            if (IGR_DAY_INTERFACE.RowCount < 1)
            {
                return;
            }

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            int vIDX_PERSON_ID = IGR_DAY_INTERFACE.GetColumnToIndex("PERSON_ID");
            int vIDX_WORK_DATE = IGR_DAY_INTERFACE.GetColumnToIndex("WORK_DATE");
            int vIDX_CORP_ID = IGR_DAY_INTERFACE.GetColumnToIndex("CORP_ID");
            string vSTATUS = "F";
            string vMESSAGE = null;
            for (int i = 0; i < IGR_DAY_INTERFACE.RowCount; i++)
            {
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_PERSON_ID", IGR_DAY_INTERFACE.GetCellValue(i, vIDX_PERSON_ID));
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_WORK_DATE", IGR_DAY_INTERFACE.GetCellValue(i, vIDX_WORK_DATE));
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("W_CORP_ID", IGR_DAY_INTERFACE.GetCellValue(i, vIDX_CORP_ID));
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_STATUS", "C");
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_CHECK_YN", "Y");
                IDC_SET_UPDATE_APPROVE.SetCommandParamValue("P_APPROVE_FLAG", pApproved_Flag);
                IDC_SET_UPDATE_APPROVE.ExecuteNonQuery();
                vSTATUS = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_STATUS"));
                vMESSAGE = iConv.ISNull(IDC_SET_UPDATE_APPROVE.GetCommandParamValue("O_MESSAGE"));
                if (IDC_SET_UPDATE_APPROVE.ExcuteError || vSTATUS == "F")
                {
                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    Application.DoEvents();
                    if (vMESSAGE != string.Empty)
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }
            }

            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            Search_DB();
        }

        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----        
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
                    if (IDA_DAY_INTERFACE_TRANS.IsFocused)
                    {
                        IDA_DAY_INTERFACE_TRANS.Update();                        
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_DAY_INTERFACE_TRANS.IsFocused)
                    {
                        IDA_DAY_INTERFACE_TRANS.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_DAY_INTERFACE_TRANS.IsFocused)
                    {
                        IDA_DAY_INTERFACE_TRANS.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----

        private void HRMF0311_Load(object sender, EventArgs e)
        {            
            W_SET_INTERFACE_FLAG.BringToFront();
        }

        private void HRMF0311_Shown(object sender, EventArgs e)
        {
            WORK_DATE_0.EditValue = DateTime.Today;

            // CORP SETTING
            DefaultCorporation();
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
            irbALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            TRANSFER_YN.EditValue = irbALL.RadioCheckedString;
            
            IDA_DAY_INTERFACE_TRANS.FillSchema();
        }

        private void irbALL_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv isINOUT = sender as ISRadioButtonAdv;
            TRANSFER_YN.EditValue = isINOUT.RadioCheckedString;

            // refill.
            Search_DB();
        }

        private void ibtSET_DAY_INTERFACE_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 출퇴근 집계

            string mSTATUS = "F";
            string mMessage = null;

            if (CORP_ID_0.EditValue == null)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (WORK_DATE_0.EditValue == null)
            {// 근무일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WORK_DATE_0.Focus();
                return;
            }

            idcSET_DAY_INTERFACE.ExecuteNonQuery();
            mSTATUS = idcSET_DAY_INTERFACE.GetCommandParamValue("O_STATUS").ToString();
            mMessage = iConv.ISNull(idcSET_DAY_INTERFACE.GetCommandParamValue("O_MESSAGE"));
            if (idcSET_DAY_INTERFACE.ExcuteError || mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            MessageBoxAdv.Show(mMessage, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // refill.
            Search_DB();
      
        }

        private void ibtTRANS_OK_ButtonClick(object pSender, EventArgs pEventArgs)
        {// 이첩처리
            string mSTATUS = "F";
            string mMessage = null;

            idcDAY_INTERFACE_TRANS.SetCommandParamValue("W_CAP_CHECK_YN", "N");
            idcDAY_INTERFACE_TRANS.ExecuteNonQuery();
            mSTATUS = idcDAY_INTERFACE_TRANS.GetCommandParamValue("O_STATUS").ToString();
            mMessage = iConv.ISNull(idcDAY_INTERFACE_TRANS.GetCommandParamValue("O_MESSAGE"));
            if (idcDAY_INTERFACE_TRANS.ExcuteError || mSTATUS == "F")
            {
                MessageBoxAdv.Show(mMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // refill.
            Search_DB();
        }

        private void BTN_SET_CHECK_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int vCNT = 0;
            foreach (DataRow vROW in IDA_DAY_INTERFACE_TRANS.CurrentRows)
            {
                if (vROW.RowState == DataRowState.Unchanged)
                {
                    //
                }
                else
                {
                    vCNT++;
                }
            }
            if (vCNT != 0)
            {
                if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10150"), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    IDA_DAY_INTERFACE_TRANS.Update();
                }
                else
                {
                    IDA_DAY_INTERFACE_TRANS.Cancel();
                }
            }
            Set_Update_Review("B"); 
        }

        private void ibtnUP_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(1);
        }

        private void ibtnDOWN_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            isSearch_Day_History(-1);
        }
        #endregion  

        #region ----- Adapter Event -----
        private void idaDAY_INTERFACE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Person ID(사원 정보)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["WORK_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Work Date(근무일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation Name(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaDAY_INTERFACE_PreDelete(ISPreDeleteEventArgs e)
        {            
        }

        private void idaDAY_INTERFACE_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            //isSearch_WorkCalendar(igrDAY_INTERFACE.GetCellValue("PERSON_ID"), igrDAY_INTERFACE.GetCellValue("WORK_DATE"));
            WORK_DATE_8.EditValue = WORK_DATE_0.EditValue;
            isSearch_Day_History(0);
        }

        #endregion

        #region ----- LookUp Event ------

        private void ILA_W_OPERATING_UNIT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ildWORK_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "WORK_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFLOOR_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildPERSON.SetLookupParamValue("W_END_DATE", WORK_DATE_0.EditValue);
        }

        private void ilaDUTY_MODIFY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "DUTY_MODIFY");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void idaYES_NO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "YES_NO");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        #endregion

    }
}