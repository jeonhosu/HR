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

namespace HRMF0141
{
    public partial class HRMF0141 : Office2007Form
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #endregion;

        #region ----- Constructor -----

        public HRMF0141()
        {
            InitializeComponent();
        }

        public HRMF0141(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void Search_DB()
        {
            if (itbINSUR.SelectedTab.TabIndex == 1)
            {
                if (iString.ISNull(HEALTH_YYYYMM_0.EditValue) == String.Empty)
                {// 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    HEALTH_YYYYMM_0.Focus();
                    return;
                }
                idaHEALTH_YYYYMM.Fill();
                idaHEALTH_INSUR.Fill();
                igrHEALTH_INSUR.Focus();
            }
            else if (itbINSUR.SelectedTab.TabIndex == 2)
            {
                if (iString.ISNull(PENSION_YYYYMM_0.EditValue) == String.Empty)
                {// 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    PENSION_YYYYMM_0.Focus();
                    return;
                }
                idaPENSION_YYYYMM.Fill();
                idaPENSION_INSUR.Fill();
                igrPENSION_INSUR.Focus();
            }
            else if (itbINSUR.SelectedTab.TabIndex == 3)
            {
                if (iString.ISNull(EMP_YYYYMM_0.EditValue) == String.Empty)
                {// 
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10036"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    EMP_YYYYMM_0.Focus();
                    return;
                }
                idaEMP_YYYYMM.Fill();
                idaEMP_INSUR.Fill();
                igrEMP_INSUR.Focus();
            }
        }

        private void SetCommonParameter(object pGroup_Code, object pEnable_YN)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", pEnable_YN);
        }

        private void Insert_Health_YYYYMM()
        {
            idaHEALTH_INSUR.Fill();
            HEALTH_YYYYMM.Focus();
        }

        private void Insert_Health_Insur()
        {
            Int32 IDX_START = igrHEALTH_INSUR.GetColumnToIndex("START_AMOUNT");
            igrHEALTH_INSUR.CurrentCellMoveTo(IDX_START);
            igrHEALTH_INSUR.CurrentCellActivate(IDX_START);
            igrHEALTH_INSUR.Focus();
        }

        private void Insert_Pension_YYYYMM()
        {
            idaPENSION_INSUR.Fill();
            PENSION_YYYYMM.Focus();
        }

        private void Insert_Pension_Insur()
        {
            Int32 IDX_START = igrPENSION_INSUR.GetColumnToIndex("START_AMOUNT");
            igrPENSION_INSUR.CurrentCellMoveTo(IDX_START);
            igrPENSION_INSUR.CurrentCellActivate(IDX_START);
            igrPENSION_INSUR.Focus();
        }

        private void Insert_EMP_YYYYMM()
        {
            idaEMP_INSUR.Fill();
            EMP_YYYYMM.Focus();
        }

        private void Insert_EMP_Insur()
        {
            Int32 IDX_START = igrEMP_INSUR.GetColumnToIndex("START_AMOUNT");
            igrEMP_INSUR.CurrentCellMoveTo(IDX_START);
            igrEMP_INSUR.CurrentCellActivate(IDX_START);
            igrEMP_INSUR.Focus();
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
                    if (idaHEALTH_YYYYMM.IsFocused)
                    {
                        idaHEALTH_YYYYMM.AddOver();
                        Insert_Health_YYYYMM();
                    }
                    else if (idaHEALTH_INSUR.IsFocused)
                    {
                        if (iString.ISNull(HEALTH_YYYYMM.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        idaHEALTH_INSUR.AddOver();
                        Insert_Health_Insur();
                    }
                    else if (idaPENSION_YYYYMM.IsFocused)
                    {
                        idaPENSION_YYYYMM.AddOver();
                        Insert_Pension_YYYYMM();
                    }
                    else if (idaPENSION_INSUR.IsFocused)
                    {
                        if (iString.ISNull(PENSION_YYYYMM.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        idaPENSION_INSUR.AddOver();
                        Insert_Pension_Insur();
                    }
                    else if (idaEMP_YYYYMM.IsFocused)
                    {
                        idaEMP_YYYYMM.AddOver();
                        Insert_EMP_YYYYMM();
                    }
                    else if (idaEMP_INSUR.IsFocused)
                    {
                        if (iString.ISNull(EMP_YYYYMM.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        idaEMP_INSUR.AddOver();
                        Insert_EMP_Insur();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaHEALTH_YYYYMM.IsFocused)
                    {
                        idaHEALTH_YYYYMM.AddUnder();
                        Insert_Health_YYYYMM();
                    }
                    else if (idaHEALTH_INSUR.IsFocused)
                    {
                        if (iString.ISNull(HEALTH_YYYYMM.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        idaHEALTH_INSUR.AddUnder();
                        Insert_Health_Insur();
                    }
                    else if (idaPENSION_YYYYMM.IsFocused)
                    {
                        idaPENSION_YYYYMM.AddUnder();
                        Insert_Pension_YYYYMM();
                    }
                    else if (idaPENSION_INSUR.IsFocused)
                    {
                        if (iString.ISNull(PENSION_YYYYMM.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        idaPENSION_INSUR.AddUnder();
                        Insert_Pension_Insur();
                    }
                    else if (idaEMP_YYYYMM.IsFocused)
                    {
                        idaEMP_YYYYMM.AddUnder();
                        Insert_EMP_YYYYMM();
                    }
                    else if (idaEMP_INSUR.IsFocused)
                    {
                        if (iString.ISNull(EMP_YYYYMM.EditValue) == string.Empty)
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        idaEMP_INSUR.AddUnder();
                        Insert_EMP_Insur();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaHEALTH_YYYYMM.IsFocused || idaHEALTH_INSUR.IsFocused)
                    {
                        idaHEALTH_YYYYMM.Update();
                        idaHEALTH_INSUR.Update();
                    }
                    else if (idaPENSION_YYYYMM.IsFocused || idaPENSION_INSUR.IsFocused)
                    {
                        idaPENSION_YYYYMM.Update();
                        idaPENSION_INSUR.Update();
                    }
                    else if (idaEMP_YYYYMM.IsFocused || idaEMP_INSUR.IsFocused)
                    {
                        idaEMP_YYYYMM.Update();
                        idaEMP_INSUR.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaHEALTH_YYYYMM.IsFocused)
                    {
                        idaHEALTH_INSUR.Cancel();
                        idaHEALTH_YYYYMM.Cancel();
                    }
                    else if (idaHEALTH_INSUR.IsFocused)
                    {
                        idaHEALTH_INSUR.Cancel();
                    }
                    else if (idaPENSION_YYYYMM.IsFocused)
                    {
                        idaPENSION_INSUR.Cancel();
                        idaPENSION_YYYYMM.Cancel();
                    }
                    else if (idaPENSION_INSUR.IsFocused)
                    {
                        idaPENSION_INSUR.Cancel();
                    }
                    else if (idaEMP_YYYYMM.IsFocused)
                    {
                        idaEMP_INSUR.Cancel();
                        idaEMP_YYYYMM.Cancel();
                    }
                    else if (idaEMP_INSUR.IsFocused)
                    {
                        idaEMP_INSUR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaHEALTH_YYYYMM.IsFocused)
                    {
                        idaHEALTH_YYYYMM.Delete();
                    }
                    else if (idaHEALTH_INSUR.IsFocused)
                    {
                        idaHEALTH_INSUR.Delete();
                    }
                    else if (idaPENSION_YYYYMM.IsFocused)
                    {
                        idaPENSION_INSUR.Delete();
                        idaPENSION_YYYYMM.Delete();
                    }
                    else if (idaPENSION_INSUR.IsFocused)
                    {
                        idaPENSION_INSUR.Delete();
                    }
                    else if (idaEMP_YYYYMM.IsFocused)
                    {
                        idaEMP_INSUR.Delete();
                        idaEMP_YYYYMM.Delete();
                    }
                    else if (idaEMP_INSUR.IsFocused)
                    {
                        idaEMP_INSUR.Delete();
                    }
                }
            }
        }

        #endregion;

        #region ----- Form Event ------

        private void HRMF0141_Load(object sender, EventArgs e)
        {
            idaHEALTH_YYYYMM.FillSchema();
            idaHEALTH_INSUR.FillSchema();
            idaPENSION_YYYYMM.FillSchema();
            idaPENSION_INSUR.FillSchema();
            idaEMP_YYYYMM.FillSchema();
            idaEMP_INSUR.FillSchema();
        }

        private void HRMF0141_Shown(object sender, EventArgs e)
        {
            itbINSUR.SelectedIndex = 0;
            itbINSUR.Focus();
        }

        private void itbINSUR_Click(object sender, EventArgs e)
        {
            if (itbINSUR.SelectedTab.TabIndex == 1)
            {
                igrHEALTH_INSUR.Focus();
            }
            else if (itbINSUR.SelectedTab.TabIndex == 2)
            {
                igrPENSION_INSUR.Focus();
            }
        }

        private void igrHEALTH_INSUR_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (igrHEALTH_INSUR.GetColumnToIndex("PERSON_RATE") == e.ColIndex)
            {
                igrHEALTH_INSUR.SetCellValue("CORPORATION_RATE", e.NewValue);
            }
        }

        private void igrPENSION_INSUR_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            if (igrPENSION_INSUR.GetColumnToIndex("PERSON_RATE") == e.ColIndex)
            {
                igrPENSION_INSUR.SetCellValue("CORPORATION_RATE", e.NewValue);
            }
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaHEALTH_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildINSUR_YYYYMM.SetLookupParamValue("P_INSUR_TYPE", INSUR_TYPE_M.EditValue);
        }

        private void ilaPENSION_YYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildINSUR_YYYYMM.SetLookupParamValue("P_INSUR_TYPE", INSUR_TYPE_P.EditValue);
        }

        private void ilaEMP_YYYYMM_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildINSUR_YYYYMM.SetLookupParamValue("P_INSUR_TYPE", INSUR_TYPE_E.EditValue);
        }

        private void ilaAMOUNT_TYPE_H_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("AMOUNT_TYPE", "Y");
        }

        private void ilaNUMBER_LEVEL_H_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("NUMBER_LEVEL_TYPE", "Y");
        }

        private void ilaNUMBER_CAL_H_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("NUMBER_CAL_TYPE", "Y");
        }

        private void ilaAMOUNT_TYPE_P_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("AMOUNT_TYPE", "Y");
        }

        private void ilaNUMBER_LEVEL_P_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("NUMBER_LEVEL_TYPE", "Y");
        }

        private void ilaNUMBER_CAL_P_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("NUMBER_CAL_TYPE", "Y");
        }

        private void ilaAMOUNT_TYPE_E_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("AMOUNT_TYPE", "Y");
        }

        private void ilaNUMBER_LEVEL_E_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("NUMBER_LEVEL_TYPE", "Y");
        }

        private void ilaNUMBER_CAL_E_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("NUMBER_CAL_TYPE", "Y");
        }

        private void ilaHEALTH_YYYYMM_PrePopupShow_1(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(), 8)));
        }

        private void ilaPENSION_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(), 8)));
        }

        private void ilaEMP_YYYYMM_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildYYYYMM.SetLookupParamValue("W_END_YYYYMM", iDate.ISYearMonth(iDate.ISDate_Month_Add(iDate.ISGetDate(), 8)));
        }

        #endregion

        #region ----- Adapter Event -----
        //건강보험
        private void idaHEALTH_YYYYMM_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["INSUR_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaHEALTH_INSUR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10373"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CORPORATION_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10374"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["AMOUNT_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10370"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["NUMBER_LEVEL_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10371"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["NUMBER_CAL_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10372"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        
        //국민연금.
        private void idaPENSION_YYYYMM_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["INSUR_YYYYMM"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10375"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        private void idaPENSION_INSUR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (iString.ISNull(e.Row["PERSON_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10373"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["CORPORATION_RATE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10374"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["AMOUNT_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10370"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["NUMBER_LEVEL_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10371"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (iString.ISNull(e.Row["NUMBER_CAL_TYPE"]) == string.Empty)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10372"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }

        #endregion

    }
}