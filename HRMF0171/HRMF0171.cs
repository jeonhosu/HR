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

namespace HRMF0171
{
    public partial class HRMF0171 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public HRMF0171(Form pMainFom, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainFom;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #region ----- Method -----
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

        }

        private void SEARCH_DB()
        {

            IDA_DEPT_MASTER.Fill();

        }

      

        private void Insert_Dept_Mapping()
        {
      

        }

        #endregion

        #region ----- main Button Click ------
        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SEARCH_DB();

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
             
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
               
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_DEPT_MASTER.Update();
                    IDA_DEPT_MASTER.Fill();
            
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                  
                }
            }
        }
        #endregion

        #region ----- Form Event -----
        private void HRMF0171_Load(object sender, EventArgs e)
        {
            IDA_DEPT_MASTER.FillSchema();
         

            DefaultCorporation();
        
        }
        #endregion

        #region ---- Adapter Event -----
        private void isDataAdapter1_PreRowUpdate(ISPreRowUpdateEventArgs e)
        {             
      
        }

        private void isDataAdapter1_PreDelete(ISPreDeleteEventArgs e)
        {
           
        }

        private void idaDEPT_MAPPING_PreRowUpdate(ISPreRowUpdateEventArgs e)
        { 
      
        }

        private void idaDEPT_MAPPING_PreDelete(ISPreDeleteEventArgs e)
        {
      
        }

        #endregion

        #region ---- Lookup Event -----

    
        private void ilaDEPT_UPPER_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
        
        }

        private void ilaMODULE_TYPE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
        }

        private void ilaMODULE_TYPE_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
         
        }
        
        private void ilaHR_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {

        }

        private void ilaM_DEPT_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
     
        }

        private void ilaM_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
           
        }

        private void ILA_PERSON_VALUER_1_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
      
        }

        private void ILA_PERSON_VALUER_2_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
       
        }

        #endregion

        private void IGR_DEPT_DETAIL_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            switch (IGR_DEPT_DETAIL.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "WEEK_DUTY":

                    IGR_DEPT_DETAIL.SetCellValue("TO_TOTAL", Convert.ToDecimal(e.NewValue) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("NIGHT_DUTY")) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("WEEK_DUTY_OUT")) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("NIGHT_DUTY_OUT")));
                    
                    break;

                case "NIGHT_DUTY":
                    IGR_DEPT_DETAIL.SetCellValue("TO_TOTAL", Convert.ToDecimal(e.NewValue) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("WEEK_DUTY")) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("WEEK_DUTY_OUT")) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("NIGHT_DUTY_OUT")));
                    break;
                
                case "WEEK_DUTY_OUT":
                    IGR_DEPT_DETAIL.SetCellValue("TO_TOTAL", Convert.ToDecimal(e.NewValue) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("NIGHT_DUTY")) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("WEEK_DUTY")) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("NIGHT_DUTY_OUT")));
                    break;

                case "NIGHT_DUTY_OUT":
                    IGR_DEPT_DETAIL.SetCellValue("TO_TOTAL", Convert.ToDecimal(e.NewValue) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("NIGHT_DUTY")) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("WEEK_DUTY_OUT")) + Convert.ToDecimal(IGR_DEPT_DETAIL.GetCellValue("WEEK_DUTY")));
                    break;

                default:
                    break;
               
            }
        }

    }
}