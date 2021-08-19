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

namespace HRMF0821
{
    public partial class HRMF0821 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;
        
        #region ----- Constructor -----
        public HRMF0821(Form pMainForm, ISAppInterface pAppInterface)
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

        private void Init_Insert()
        {
            igrFOOD_COUPON.SetCellValue("FOOD_DATE", START_DATE_0.EditValue);
            igrFOOD_COUPON.SetCellValue("DEVICE_ID", DEVICE_ID_0.EditValue);
            igrFOOD_COUPON.SetCellValue("DEVICE_NAME", DEVICE_NAME_0.EditValue);
            igrFOOD_COUPON.SetCellValue("CORP_ID", CORP_ID_0.EditValue);
            igrFOOD_COUPON.SetCellValue("CORP_NAME", CORP_NAME_0.EditValue);

            igrFOOD_COUPON.SetCellValue("FOOD_1_COUNT", 0);
            igrFOOD_COUPON.SetCellValue("FOOD_2_COUNT", 0);
            igrFOOD_COUPON.SetCellValue("FOOD_3_COUNT", 0);
            igrFOOD_COUPON.SetCellValue("FOOD_4_COUNT", 0);
            igrFOOD_COUPON.SetCellValue("SNACK_1_COUNT", 0);
            igrFOOD_COUPON.SetCellValue("SNACK_2_COUNT", 0);
            igrFOOD_COUPON.SetCellValue("SNACK_3_COUNT", 0);
            igrFOOD_COUPON.SetCellValue("SNACK_4_COUNT", 0);
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

            idaFOOD_COUPON.Fill();
            igrFOOD_COUPON.Focus();
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
                    if(idaFOOD_COUPON.IsFocused)
                    {
                        idaFOOD_COUPON.AddOver();                        
                    }
                    Init_Insert();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaFOOD_COUPON.IsFocused)
                    {
                        idaFOOD_COUPON.AddUnder();
                    }
                    Init_Insert();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaFOOD_COUPON.IsFocused)
                    {
                        idaFOOD_COUPON.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaFOOD_COUPON.IsFocused)
                    {
                        idaFOOD_COUPON.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaFOOD_COUPON.IsFocused)
                    {
                        idaFOOD_COUPON.Delete();
                    }
                }
            }
        }
        #endregion;

        #region ----- Form Event -----
        private void HRMF0821_Load(object sender, EventArgs e)
        {
            ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

            idaFOOD_COUPON.FillSchema();

            START_DATE_0.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE_0.EditValue = iDate.ISGetDate();

            //DefaultSetFormReSize();		//[Child Form, Mdi Form에 맞게 ReSize]           
        }
        #endregion  

        #region ----- Adapter Event -----
        private void idaFOOD_COUPON_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {
            if (e.Row["DEVICE_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Cafeteria Name(식당)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }
            if (e.Row["CORP_ID"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Corporation Name(업체)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }            
            if (e.Row["FOOD_DATE"] == DBNull.Value)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10037", "&&VALUE:=Food Date(식사일자)"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        private void ilaCORP_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }
        private void ilaCAFETERIA_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCAFETERIA.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        private void ilaCORP_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");
        }
        #endregion
                
    }
}