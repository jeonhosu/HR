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

namespace HRMF0502
{
    public partial class HRMF0502 : Office2007Form
    {
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        EAPF1102.EAPF1102 mEAPF1102 = new EAPF1102.EAPF1102();

        Object mSESSION_ID;

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF0502(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
            if (iConv.ISNull(isAppInterfaceAdv1.AppInterface.Attribute_A) != string.Empty)
            {
                CORP_TYPE.EditValue = isAppInterfaceAdv1.AppInterface.Attribute_A;
            }
        }

        #endregion;

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
            ildCORP.SetLookupParamValue("W_PAY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_PAY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();

            CORP_NAME_0.BringToFront();
            igbCORP_GROUP_0.BringToFront();
            igbCORP_GROUP_0.Visible = false;

            if (iConv.ISNull(CORP_TYPE.EditValue) == "ALL")
            {
                igbCORP_GROUP_0.Visible = true; //.Show();
                igbCORP_GROUP_0.BringToFront();

                irb_ALL_0.RadioButtonValue = "A";
                CORP_TYPE.EditValue = "A";
                CORP_TYPE.BringToFront();
            }
            else if (iConv.ISNull(CORP_TYPE.EditValue) == "1")
            {
                CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
                CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");
            }

            IDC_GET_SESSION_ID_P.ExecuteNonQuery();
            mSESSION_ID = IDC_GET_SESSION_ID_P.GetCommandParamValue("O_SESSION_ID");
        }

        private void Initial_Insert_Pay_Master()
        {
            // Pay Master Initial....
            START_YYYYMM.EditValue = STD_YYYYMM_0.EditValue;
            CORP_ID.EditValue = igrPERSON.GetCellValue("CORP_ID");
            PERSON_ID.EditValue = igrPERSON.GetCellValue("PERSON_ID");
            PAY_GRADE_NAME.EditValue = igrPERSON.GetCellValue("PAY_GRADE_NAME");
            PAY_GRADE_ID.EditValue = igrPERSON.GetCellValue("PAY_GRADE_ID");
            PAY_PROVIDE_YN.CheckBoxValue = "Y";
            BONUS_PROVIDE_YN.CheckBoxValue = "Y";
            YEAR_PROVIDE_YN.CheckBoxValue = "Y";
            HIRE_INSUR_YN.CheckBoxValue = "Y";

            HEADER_DATA_STATE.EditValue = "U";
            INCOME_TAX_RATE.EditValue = 100;
            PRINT_TYPE_NAME.Focus();
        }

        private void Initial_Insert_Allowance()
        {
            igrPAY_ALLOWANCE.SetCellValue("ENABLED_FLAG", "Y");
        }

        private void Initial_Insert_Deduction()
        {
            igrPAY_DEDUCTION.SetCellValue("ENABLED_FLAG", "Y");
        }

        private void Search_DB()
        {
            if (CORP_ID_0.EditValue == null&& CORP_TYPE.EditValue.ToString() != "4")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }

            if (iConv.ISNull(STD_YYYYMM_0.EditValue) == String.Empty)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                STD_YYYYMM_0.Focus();
                return;
            }

            string vPERSON_NAME = iConv.ISNull(igrPERSON.GetCellValue("NAME"));
            int vIDX_Col = igrPERSON.GetColumnToIndex("NAME");
            idaPERSON.Fill();
            if (igrPERSON.RowCount > 0)
            {
                for (int vRow = 0; vRow < igrPERSON.RowCount; vRow++)
                {
                    if (vPERSON_NAME == iConv.ISNull(igrPERSON.GetCellValue(vRow, vIDX_Col)))
                    {
                        igrPERSON.CurrentCellActivate(vRow, vIDX_Col);
                        igrPERSON.CurrentCellMoveTo(vRow, vIDX_Col);
                    }
                }
            }
            igrPERSON.Focus();
        }

        private void isSET_ALLOWANCE()
        {
            idaGRADE_STEP_AMOUNT.Fill();

            if (idaGRADE_STEP_AMOUNT.OraSelectData.Rows.Count == 0)
            {
                return;
            }
            
            // 반환된 RECORD COUNT 만큼 루프를 돌며 GRID의 값과 비교 --> 있으면 수정, 없으면 ADD.
            for (int R = 0; R < idaGRADE_STEP_AMOUNT.OraSelectData.Rows.Count; R++)
            {
                int Row_Index = -1;
                int RECORD_VALUE = Convert.ToInt32(idaGRADE_STEP_AMOUNT.OraSelectData.Rows[R]["ALLOWANCE_ID"]);

                for (int GR = 0; GR < igrPAY_ALLOWANCE.RowCount; GR++)
                {
                    int GRID_VALUE = Convert.ToInt32(igrPAY_ALLOWANCE.GetCellValue(GR, igrPAY_ALLOWANCE.GetColumnToIndex("ALLOWANCE_ID")));
                    if (RECORD_VALUE == GRID_VALUE)
                    {
                        Row_Index = GR;
                    }
                }
                if (Row_Index == -1)
                {
                    idaPAY_ALLOWANCE.AddUnder();
                    igrPAY_ALLOWANCE.SetCellValue("PAY_HEADER_ID", PAY_HEADER_ID.EditValue);
                    igrPAY_ALLOWANCE.SetCellValue("ALLOWANCE_ID", idaGRADE_STEP_AMOUNT.OraSelectData.Rows[R]["ALLOWANCE_ID"]);
                    igrPAY_ALLOWANCE.SetCellValue("ALLOWANCE_NAME", idaGRADE_STEP_AMOUNT.OraSelectData.Rows[R]["ALLOWANCE_NAME"]);
                    igrPAY_ALLOWANCE.SetCellValue("ALLOWANCE_AMOUNT", idaGRADE_STEP_AMOUNT.OraSelectData.Rows[R]["ALLOWANCE_AMOUNT"]);
                    igrPAY_ALLOWANCE.SetCellValue("ENABLED_FLAG", "Y");
                }
                else
                {
                    igrPAY_ALLOWANCE.SetCellValue(Row_Index, igrPAY_ALLOWANCE.GetColumnToIndex("ALLOWANCE_AMOUNT"), idaGRADE_STEP_AMOUNT.OraSelectData.Rows[R]["ALLOWANCE_AMOUNT"]);
                }
            }
        }

        private void isSet_Adapter_Status()
        {
            for (int hr = 0; hr < idaPAY_MASTER_HEADER.SelectRows.Count; hr++)
            {
                if (idaPAY_MASTER_HEADER.SelectRows[hr].RowState != DataRowState.Unchanged)
                {
                    for (int r = 0; r < igrPAY_ALLOWANCE.RowCount; r++)
                    {//지급사항 수정.
                        idaPAY_ALLOWANCE.SelectRows[r].AcceptChanges();
                        idaPAY_ALLOWANCE.SelectRows[r].SetAdded();
                    }

                    for (int r = 0; r < igrPAY_DEDUCTION.RowCount ; r++)
                    {//공제사항 수정.
                        idaPAY_DEDUCTION.SelectRows[r].AcceptChanges();
                        idaPAY_DEDUCTION.SelectRows[r].SetAdded();
                    }
                }
            }
        }

        private void isSet_HEADER_Status()
        {
            idaPERSON.MoveFirst("");
            igrPERSON.BeginUpdate();
            for (int m = 0; m < idaPERSON.OraSelectData.Rows.Count; m++)
            {
                for (int i = 0; i < idaPAY_MASTER_HEADER.CurrentRows.Count; i++)
                {
                    if (idaPAY_MASTER_HEADER.CurrentRows[i].RowState == DataRowState.Unchanged)
                    {
                        for (int j = 0; j < idaPAY_ALLOWANCE.CurrentRows.Count; j++)
                        {
                            if (idaPAY_ALLOWANCE.CurrentRows[j].RowState != DataRowState.Unchanged)
                            {
                                if (idaPAY_MASTER_HEADER.CurrentRows[i].RowState == DataRowState.Unchanged)
                                {
                                    START_YYYYMM.EditValue = STD_YYYYMM_0.EditValue;
                                }
                            }
                        }
                        idaPAY_MASTER_HEADER.MoveNext("");
                    }
                }
                idaPERSON.MoveNext("");
            }
            igrPERSON.EndUpdate();

            //foreach (DataRow row in idaPAY_ALLOWANCE.OraSelectData.Rows)
            //{
            //    if (row.RowState != DataRowState.Unchanged)
            //    {
            //        object vob = row["MasterKeyId"];
            //        idaPAY_MASTER_HEADER.OraSelectData.Rows[(int)row["MasterKeyId"]]["START_YYYYMM"] = STD_YYYYMM_0.EditValue;
            //    }
            //}

            //decimal vPAY_HEADER_ID = 0;
            // 지급항목 변경 여부 체크 하여 헤더 상태 변경//
            //for (int vROW = 0; vROW < idaPAY_ALLOWANCE.SelectRows.Count; vROW++)
            //{
            //    if (idaPAY_ALLOWANCE.SelectRows[vROW].RowState != DataRowState.Unchanged)
            //    {
            //        if (idaPAY_ALLOWANCE.MasterAdapter.CurrentRow.RowState == DataRowState.Unchanged)
            //        {
            //            idaPAY_ALLOWANCE.MasterAdapter.CurrentRow["START_YYYYMM"] = STD_YYYYMM_0.EditValue;
            //        }

            //        //vPAY_HEADER_ID = iString.ISDecimaltoZero(idaPAY_ALLOWANCE.SelectRows[vROW]["PAY_HEADER_ID"]);
            //        //idaPAY_MASTER_HEADER.MoveFirst(START_YYYYMM.Name);
            //        //for (int r = 0; r < idaPAY_MASTER_HEADER.SelectRows.Count; r++)
            //        //{
            //        //    if (iString.ISDecimaltoZero(idaPAY_MASTER_HEADER.SelectRows[r]["PAY_HEADER_ID"]) == vPAY_HEADER_ID)
            //        //    {
            //        //        if (idaPAY_MASTER_HEADER.SelectRows[r].RowState == DataRowState.Unchanged)
            //        //        {
            //        //            //위치 이동(아답터)
            //        //            idaPAY_ALLOWANCE.MasterAdapter.CurrentRow["START_YYYYMM"] = STD_YYYYMM_0.EditValue;
            //        //            idaPAY_ALLOWANCE.MasterAdapter.CurrentRow.SetModified();
            //        //        }
            //        //    }
            //        //}
            //    }
            //}

            //// 공제항목 변경 여부 체크 하여 헤더 상태 변경//
            //for (int vROW = 0; vROW < idaPAY_DEDUCTION.SelectRows.Count; vROW++)
            //{
            //    if (idaPAY_DEDUCTION.SelectRows[vROW].RowState != DataRowState.Unchanged)
            //    {
            //        vPAY_HEADER_ID = iString.ISDecimaltoZero(idaPAY_DEDUCTION.SelectRows[vROW]["PAY_HEADER_ID"]);
            //        for (int r = 0; r < idaPAY_MASTER_HEADER.SelectRows.Count; r++)
            //        {
            //            if (iString.ISDecimaltoZero(idaPAY_MASTER_HEADER.SelectRows[r]["PAY_HEADER_ID"]) == vPAY_HEADER_ID)
            //            {
            //                if (idaPAY_MASTER_HEADER.SelectRows[r].RowState == DataRowState.Unchanged)
            //                {
            //                    idaPAY_ALLOWANCE.MasterAdapter.CurrentRow["START_YYYYMM"] = STD_YYYYMM_0.EditValue;
            //                    idaPAY_ALLOWANCE.MasterAdapter.CurrentRow.SetModified();
            //                }
            //            }
            //        }
            //    }
            //}
        }

        private void Init_Sum_Allowance_Amount()
        {
            decimal vSum_Amount = 0;
            //지급
            foreach (System.Data.DataRow vRow in idaPAY_ALLOWANCE.CurrentRows)
            {
                if (iConv.ISNull(vRow["ENABLED_FLAG"]) == "Y")
                {
                    vSum_Amount = vSum_Amount + iConv.ISDecimaltoZero(vRow["ALLOWANCE_AMOUNT"]);
                }
            }
            SUM_ALLOWANCE_AMOUNT.EditValue = vSum_Amount; 
        }

        private void Init_Sum_Deduction_Amount()
        {
            decimal vSum_Amount = 0;
            //공제
            foreach (System.Data.DataRow vRow in idaPAY_DEDUCTION.CurrentRows)
            {
                if (iConv.ISNull(vRow["ENABLED_FLAG"]) == "Y")
                {
                    vSum_Amount = vSum_Amount + iConv.ISDecimaltoZero(vRow["ALLOWANCE_AMOUNT"]);
                }
            }
            SUM_DEDUCTION_AMOUNT.EditValue = vSum_Amount;
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

        private object Get_Grid_Prompt(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pCol_Index)
        {
            int mCol_Count = pGrid.GridAdvExColElement[pCol_Index].HeaderElement.Count;
            object mPrompt = null;
            switch (isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage)
            {
                case ISUtil.Enum.TerritoryLanguage.Default:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].Default;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL1_KR:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL1_KR;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL2_CN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL2_CN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL3_VN:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL3_VN;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL4_JP:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL4_JP;
                        }
                    }
                    break;
                case ISUtil.Enum.TerritoryLanguage.TL5_XAA:
                    for (int r = 0; r < mCol_Count; r++)
                    {
                        if (iConv.ISNull(pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA) != string.Empty)
                        {
                            mPrompt = pGrid.GridAdvExColElement[pCol_Index].HeaderElement[r].TL5_XAA;
                        }
                    }
                    break;
            }
            return mPrompt;
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
                    if (idaPERSON.IsFocused || idaPAY_MASTER_HEADER.IsFocused)
                    {
                        idaPAY_MASTER_HEADER.AddOver();
                        Initial_Insert_Pay_Master();
                    }
                    else if (idaPAY_ALLOWANCE.IsFocused)
                    {
                        idaPAY_ALLOWANCE.AddOver();
                        Initial_Insert_Allowance();
                    }
                    else if (idaPAY_DEDUCTION.IsFocused)
                    {
                        idaPAY_DEDUCTION.AddOver();
                        Initial_Insert_Deduction();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaPERSON.IsFocused || idaPAY_MASTER_HEADER.IsFocused)
                    {
                        idaPAY_MASTER_HEADER.AddUnder();
                        Initial_Insert_Pay_Master();
                    }
                    else if (idaPAY_ALLOWANCE.IsFocused)
                    {
                        idaPAY_ALLOWANCE.AddUnder();
                        Initial_Insert_Allowance();
                    }
                    else if (idaPAY_DEDUCTION.IsFocused)
                    {
                        idaPAY_DEDUCTION.AddUnder();
                        Initial_Insert_Deduction();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if(MessageBoxAdv.Show(isMessageAdapter1.ReturnText("SVEAPP_10229", string.Format("&&PERIOD_NAME:={0}", STD_YYYYMM_0.EditValue)), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return;
                    }

                    try
                    {
                        idaPERSON.Update();
                    }
                    catch
                    {
                        return;
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaPERSON.IsFocused || idaPAY_MASTER_HEADER.IsFocused)
                    {
                        HEADER_DATA_STATE.EditValue = "N";
                        idaPAY_MASTER_HEADER.Cancel();
                    }
                    else if (idaPAY_ALLOWANCE.IsFocused)
                    {                        
                        idaPAY_ALLOWANCE.Cancel();
                        int vCOUNT = 0;
                        for (int i = 0; i < idaPAY_DEDUCTION.CurrentRows.Count; i++)
                        {
                            if (idaPAY_DEDUCTION.CurrentRows[i].RowState != DataRowState.Unchanged)
                            {
                                vCOUNT = vCOUNT + 1;
                            }
                        }
                        if (vCOUNT == 0 && iConv.ISNull(HEADER_DATA_STATE.EditValue, "N") == "N")
                        {
                            idaPAY_MASTER_HEADER.Cancel();
                        }
                    }
                    else if (idaPAY_DEDUCTION.IsFocused)
                    {
                        idaPAY_DEDUCTION.Cancel();
                        int vCOUNT = 0;
                        for (int i = 0; i < idaPAY_ALLOWANCE.CurrentRows.Count; i++)
                        {
                            if (idaPAY_ALLOWANCE.CurrentRows[i].RowState != DataRowState.Unchanged)
                            {
                                vCOUNT = vCOUNT + 1;
                            }
                        }
                        if (vCOUNT == 0 && iConv.ISNull(HEADER_DATA_STATE.EditValue, "N") == "N")
                        {
                            idaPAY_MASTER_HEADER.Cancel();
                        }
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaPERSON.IsFocused || idaPAY_MASTER_HEADER.IsFocused)
                    {
                        if (idaPAY_ALLOWANCE.CurrentRow == null)
                        {

                        }
                        else if (idaPAY_ALLOWANCE.CurrentRow.RowState == DataRowState.Added)
                        {
                            HEADER_DATA_STATE.EditValue = "N";
                        }
                        else
                        {
                            HEADER_DATA_STATE.EditValue = "U";
                        }
                        if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10525"), "Delete Qeustion", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                        IDC_PAY_HEADER_DELETE.SetCommandParamValue("P_PAY_HEADER_ID", PAY_HEADER_ID.EditValue);
                        IDC_PAY_HEADER_DELETE.ExecuteNonQuery();
                        string vSTATUS = iConv.ISNull(IDC_PAY_HEADER_DELETE.GetCommandParamValue("O_STATUS"));
                        string vMESSAGE = iConv.ISNull(IDC_PAY_HEADER_DELETE.GetCommandParamValue("O_MESSAGE"));
                        if (vSTATUS == "F")
                        {
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                        else
                        {
                            if (vMESSAGE != string.Empty)
                            {
                                MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            Search_DB();
                        }
                    }
                    else if (idaPAY_ALLOWANCE.IsFocused)
                    {
                        if (idaPAY_ALLOWANCE.CurrentRow.RowState == DataRowState.Added)
                        {
                            int vCOUNT = 0;
                            for (int i = 0; i < idaPAY_ALLOWANCE.CurrentRows.Count; i++)
                            {
                                if (idaPAY_ALLOWANCE.CurrentRows[i].RowState != DataRowState.Unchanged)
                                {
                                    vCOUNT = vCOUNT + 1;
                                }
                            }
                            for (int i = 0; i < idaPAY_DEDUCTION.CurrentRows.Count; i++)
                            {
                                if (idaPAY_DEDUCTION.CurrentRows[i].RowState != DataRowState.Unchanged)
                                {
                                    vCOUNT = vCOUNT + 1;
                                }
                            }
                            if (vCOUNT == 0 && iConv.ISNull(HEADER_DATA_STATE.EditValue, "N") == "N")
                            {
                                idaPAY_MASTER_HEADER.Cancel();
                            }
                        }
                        else
                        {
                            LINE_DATA_STATE.EditValue = "U";
                        }
                        idaPAY_ALLOWANCE.Delete();                        
                    }
                    else if (idaPAY_DEDUCTION.IsFocused)
                    {
                        if (idaPAY_DEDUCTION.CurrentRow.RowState == DataRowState.Added)
                        {
                            int vCOUNT = 0;
                            for (int i = 0; i < idaPAY_ALLOWANCE.CurrentRows.Count; i++)
                            {
                                if (idaPAY_ALLOWANCE.CurrentRows[i].RowState != DataRowState.Unchanged)
                                {
                                    vCOUNT = vCOUNT + 1;
                                }
                            }
                            for (int i = 0; i < idaPAY_DEDUCTION.CurrentRows.Count; i++)
                            {
                                if (idaPAY_DEDUCTION.CurrentRows[i].RowState != DataRowState.Unchanged)
                                {
                                    vCOUNT = vCOUNT + 1;
                                }
                            }
                            if (vCOUNT == 0 && iConv.ISNull(HEADER_DATA_STATE.EditValue, "N") == "N")
                            {
                                idaPAY_MASTER_HEADER.Cancel();
                            }
                        }
                        else
                        {
                            LINE_DATA_STATE.EditValue = "U";
                        }
                        idaPAY_DEDUCTION.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void HRMF0502_Load(object sender, EventArgs e)
        {
            idaPERSON.FillSchema();
            idaPAY_MASTER_HEADER.FillSchema();
            idaPAY_ALLOWANCE.FillSchema();
            idaPAY_DEDUCTION.FillSchema();

            STD_YYYYMM_0.EditValue = iDate.ISYearMonth(DateTime.Today);
            
            DefaultCorporation();              //Default Corp.
            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]           

            SUM_ALLOWANCE_AMOUNT.BringToFront();
            SUM_DEDUCTION_AMOUNT.BringToFront();
        }

        private void irb_ALL_0_Click(object sender, EventArgs e)
        {
            ISRadioButtonAdv RB_STATUS = sender as ISRadioButtonAdv;
            CORP_TYPE.EditValue = RB_STATUS.RadioCheckedString;
        }

        private void igrPAY_ALLOWANCE_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            LINE_DATA_STATE.EditValue = "U";
        }
        
        private void igrPAY_DEDUCTION_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            LINE_DATA_STATE.EditValue = "U";
        }

        private void BTN_GEN_HOURLY_DETAIL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iConv.ISNull(STD_YYYYMM_0.EditValue) == String.Empty)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);               
                return;
            }

            if (iConv.ISNull(PERSON_ID.EditValue) == String.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10028"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                return;
            }

            HRMF0502_DETAIL vHRMF0502_DETAIL = new HRMF0502_DETAIL(this.MdiParent, isAppInterfaceAdv1.AppInterface,
                                                                    STD_YYYYMM_0.EditValue, PERSON_NAME.EditValue, PERSON_ID.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0502_DETAIL, isAppInterfaceAdv1.AppInterface);
            vHRMF0502_DETAIL.ShowDialog();
            vHRMF0502_DETAIL.Dispose();
        }

        private void BTN_EXCEL_EXPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0502_EXPORT vHRMF0502_EXPORT = new HRMF0502_EXPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface
                                                                , CORP_ID_0.EditValue, CORP_NAME_0.EditValue
                                                                , STD_YYYYMM_0.EditValue, OPERATING_UNIT_ID_0.EditValue, OPERATING_UNIT_NAME_0.EditValue
                                                                , PAY_TYPE_0.EditValue, PAY_TYPE_NAME_0.EditValue
                                                                , DEPT_ID_0.EditValue, DEPT_NAME_0.EditValue
                                                                , PAY_GRADE_ID_0.EditValue, PAY_GRADE_NAME_0.EditValue
                                                                , PERSON_ID_0.EditValue, PERSON_NUM_0.EditValue, PERSON_NAME_0.EditValue);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0502_EXPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0502_EXPORT.ShowDialog();
            vHRMF0502_EXPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        private void BTN_EXCEL_IMPORT_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            DialogResult vdlgResult;
            HRMF0502_IMPORT vHRMF0502_IMPORT = new HRMF0502_IMPORT(this.MdiParent, isAppInterfaceAdv1.AppInterface, CORP_ID_0.EditValue
                                                                , STD_YYYYMM_0.EditValue, mSESSION_ID);
            mEAPF1102.SetProperties(EAPF1102.INIT_TYPE.None, vHRMF0502_IMPORT, isAppInterfaceAdv1.AppInterface);
            vdlgResult = vHRMF0502_IMPORT.ShowDialog();
            vHRMF0502_IMPORT.Dispose();
            if (vdlgResult == DialogResult.OK)
            {
                Search_DB();
            }
        }

        #endregion  

        #region ----- Adapter Event -----

        private void idaPERSON_UpdateCompleted(object pSender)
        {
            Search_DB();
        }

        private void idaPAY_ALLOWANCE_FillCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            SUM_ALLOWANCE_AMOUNT.EditValue = 0;
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Init_Sum_Allowance_Amount();
        }

        private void idaPAY_ALLOWANCE_FilterCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            SUM_ALLOWANCE_AMOUNT.EditValue = 0;
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Init_Sum_Allowance_Amount();
        }

        private void idaPAY_DEDUCTION_FillCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            SUM_DEDUCTION_AMOUNT.EditValue = 0;
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Init_Sum_Deduction_Amount();
        }

        private void idaPAY_DEDUCTION_FilterCompleted(object pSender, DataView pOraDataView, ISBindingEventArgs pBindingManager)
        {
            SUM_DEDUCTION_AMOUNT.EditValue = 0;
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Init_Sum_Deduction_Amount();
        }


        // Pay Master 항목.
        private void idaPAY_MASTER_HEADER_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iConv.ISNull(STD_YYYYMM_0.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(STD_YYYYMM_0))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iConv.ISNull("START_YYYYMM") == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(START_YYYYMM))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PRINT_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PERSON_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PRINT_TYPE"] == DBNull.Value)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PRINT_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (e.Row["PAY_TYPE"] == DBNull.Value)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAY_TYPE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (e.Row["PAY_GRADE_ID"] == DBNull.Value)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}", Get_Edit_Prompt(PAY_GRADE_NAME))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (e.Row["CURRENCY_CODE"] == DBNull.Value)
            {
                e.Cancel = true;
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90004", String.Format("&&FIELD_NAME:={0}",Get_Edit_Prompt(CURRENCY_DESC))), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void idaPAY_MASTER_HEADER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        // Allowance 항목.
        private void idaPAY_ALLOWANCE_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["ALLOWANCE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance(항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ALLOWANCE_AMOUNT"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance Amount(금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaPAY_ALLOWANCE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        // Deduction 항목.
        private void idaPAY_DEDUCTION_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["ALLOWANCE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance(항목)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["ALLOWANCE_AMOUNT"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Allowance Amount(금액)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        private void idaPAY_DEDUCTION_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Data(해당 자료)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        #endregion

        #region ----- LookUp Event -----

        private void ilaYYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_START_YYYYMM", "2001-01");
        }

        private void ilaPAY_GRADE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_GRADE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaPAY_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ILA_OPERATING_UNIT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_OPERATING_UNIT.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }


        private void ilaDEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildDEPT.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }

        private void ilaPERSON_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            string vSTD_YYYYMM = iConv.ISNull(STD_YYYYMM_0.EditValue);
            ildPERSON_0.SetLookupParamValue("W_START_DATE", iDate.ISMonth_1st(vSTD_YYYYMM));
            ildPERSON_0.SetLookupParamValue("W_END_DATE", iDate.ISMonth_Last(vSTD_YYYYMM));
        }

        private void ilaPAY_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PAY_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaALLOWANCE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ALLOWANCE.SetLookupParamValue("W_ENABLED_FLAG", "Y");             
        }

        private void ilaDEDUCTION_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEDUCTION.SetLookupParamValue("W_ENABLED_FLAG", "Y"); 
        }

        private void ilaPRINT_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "PRINT_TYPE");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaBANK_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "BANK");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaGRADE_STEP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            if (PAY_GRADE_ID.EditValue == null)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Grade(직급)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
            if (iConv.ISNull(PAY_TYPE.EditValue) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Pay Type(급여제)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
            
            ildGRADE_STEP.SetLookupParamValue("W_ENABLED_FLAG", "Y");
        }

        private void ilaGRADE_STEP_SelectedRowData(object pSender)
        {
            isSET_ALLOWANCE();
        }

        private void ILA_CURRENCY_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_CURRENCY.SetLookupParamValue("W_EXCEPT_BASE_YN", "N");
            ILD_CURRENCY.SetLookupParamValue("W_ENABLED_YN", "Y"); 
        }

        #endregion

    }
}