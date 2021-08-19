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

namespace HRMF0802
{
    public partial class HRMF0802 : Office2007Form
    {
        #region ----- Variables -----



        #endregion;
        
        #region ----- Constructor -----
        public HRMF0802(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }
        #endregion;

        #region ----- Private Methods ----
        private void DefaultSetFormReSize()
        {//[Child Form, Mdi Form에 맞게 ReSize]
            int vMinusWidth = 4;
            int vMinusHeight = 54;
            System.Drawing.Size vSize = this.MdiParent.ClientSize;
            this.Width = vSize.Width - vMinusWidth;
            this.Height = vSize.Height - vMinusHeight;
        }

        private void isSearch_DB()
        {
            if (START_DATE_0.EditValue == null)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (END_DATE_0.EditValue == null)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }
            if (Convert.ToDateTime(START_DATE_0.EditValue) > Convert.ToDateTime(END_DATE_0.EditValue))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }

            // ACTIVE TABPACE의 GRIDE에 FOCUS 이동하기.
            if (itbMISTAKE_CHECK.SelectedTab.TabIndex == 1)
            {
                idaMISTAKE_CHECK_PERSON.Fill();
                igrPERSON.Focus();
            }
            else if (itbMISTAKE_CHECK.SelectedTab.TabIndex == 2)
            {
                idaMISTAKE_CHECK_VISITOR.Fill();
                igrVISITOR.Focus();
            }           
        }
        #endregion;

        #region ----- isAppInterfaceAdv1_AppMainButtonClick Events -----        
        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    isSearch_DB();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if(idaMISTAKE_CHECK_PERSON.IsFocused)
                    {
                        idaMISTAKE_CHECK_PERSON.AddOver();

                        igrPERSON.SetCellValue("DEVICE_ID", DEVICE_ID_0.EditValue);
                        igrPERSON.SetCellValue("DEVICE_NAME", DEVICE_NAME_0.EditValue);
                        igrPERSON.SetCellValue("FOOD_DATE", END_DATE_0.EditValue);
                    }
                    else if (idaMISTAKE_CHECK_VISITOR.IsFocused)
                    {
                        idaMISTAKE_CHECK_VISITOR.AddOver();

                        igrVISITOR.SetCellValue("DEVICE_ID", DEVICE_ID_0.EditValue);
                        igrVISITOR.SetCellValue("DEVICE_NAME", DEVICE_NAME_0.EditValue);
                        igrVISITOR.SetCellValue("FOOD_DATE", END_DATE_0.EditValue);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaMISTAKE_CHECK_PERSON.IsFocused)
                    {
                        idaMISTAKE_CHECK_PERSON.AddUnder();
                        igrPERSON.SetCellValue("DEVICE_ID", DEVICE_ID_0.EditValue);
                        igrPERSON.SetCellValue("DEVICE_NAME", DEVICE_NAME_0.EditValue);
                        igrPERSON.SetCellValue("FOOD_DATE", END_DATE_0.EditValue);
                    }
                    else if (idaMISTAKE_CHECK_VISITOR.IsFocused)
                    {
                        idaMISTAKE_CHECK_VISITOR.AddUnder();

                        igrVISITOR.SetCellValue("DEVICE_ID", DEVICE_ID_0.EditValue);
                        igrVISITOR.SetCellValue("DEVICE_NAME", DEVICE_NAME_0.EditValue);
                        igrVISITOR.SetCellValue("FOOD_DATE", END_DATE_0.EditValue);
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaMISTAKE_CHECK_PERSON.IsFocused)
                    {
                        idaMISTAKE_CHECK_PERSON.Update();
                    }
                    else if (idaMISTAKE_CHECK_VISITOR.IsFocused)
                    {
                        idaMISTAKE_CHECK_VISITOR.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaMISTAKE_CHECK_PERSON.IsFocused)
                    {
                        idaMISTAKE_CHECK_PERSON.Cancel();
                    }
                    else if (idaMISTAKE_CHECK_VISITOR.IsFocused)
                    {
                        idaMISTAKE_CHECK_VISITOR.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaMISTAKE_CHECK_PERSON.IsFocused)
                    {
                        idaMISTAKE_CHECK_PERSON.Delete();
                    }
                    else if (idaMISTAKE_CHECK_VISITOR.IsFocused)
                    {
                        idaMISTAKE_CHECK_VISITOR.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----
        private void HRMF0802_Load(object sender, EventArgs e)
        {
            ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

            idaMISTAKE_CHECK_PERSON.FillSchema();
            idaMISTAKE_CHECK_VISITOR.FillSchema();

            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = iDate.ISGetDate();

            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]           
        }
        #endregion  

        #region ----- Adapter Event -----
        private void idaMISTAKE_CHECK_PERSON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["DEVICE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10065"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10016"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["FOOD_FLAG"] != DBNull.Value && string.IsNullOrEmpty(e.Row["FOOD_FLAG"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10066"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["FOOD_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10067"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        
        private void idaMISTAKE_CHECK_VISITOR_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["DEVICE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10065"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["PERSON_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10068"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["FOOD_FLAG"] != DBNull.Value && string.IsNullOrEmpty(e.Row["FOOD_FLAG"].ToString()))
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10066"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["FOOD_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10067"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
        }
        #endregion

        #region ----- LookUp Event -----
        private void ilaDEVICE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCAFETERIA.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void ilaFOOD_FLAG_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FOOD_FLAG");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaFOOD_FLAG_VISITOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCOMMON.SetLookupParamValue("W_GROUP_CODE", "FOOD_FLAG");
            ildCOMMON.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }

        private void ilaVISITOR_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildVISITOR.SetLookupParamValue("W_USABLE_CHECK_YN", "Y");
        }
        #endregion


        
    }
}