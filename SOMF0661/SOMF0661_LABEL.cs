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

namespace SOMF0661
{
    public partial class SOMF0661_LABEL : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        
        #region ----- Variables -----


        #endregion;
        
        #region ----- Constructor -----

        public SOMF0661_LABEL(Form pMainForm, ISAppInterface pAppInterface, object pDELIVERY_ORDER_ID, object pSHIP_TO_CUST_SITE_NAME)
        {
            InitializeComponent();

            isAppInterfaceAdv1.AppInterface = pAppInterface;

            V_DELIVERY_ORDER_ID.EditValue = pDELIVERY_ORDER_ID;
            V_SHIP_TO_CUST_STIE_NAME.EditValue = pSHIP_TO_CUST_SITE_NAME;
        }

        #endregion;

        #region ----- Events -----

        private void SOMF0661_LABEL_Load(object sender, EventArgs e)
        {
            IDA_LIST.FillSchema();

            V_BARCODE.Focus();
        }

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {

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

        private void V_BARCODE_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                IDA_TARGET.SetSelectParamValue("W_BOX_NO", V_BARCODE.EditValue);
                IDA_TARGET.SetSelectParamValue("W_OUT_BOX_FLAG", "N");

                IDA_TARGET.Fill();

                decimal vTotalQty = 0;

                foreach (DataRow row in IDA_TARGET.SelectRows)
                {
                    for (int i = 0; i < ISG_LIST.RowCount; i++)
                    {
                        if (Convert.ToString(row["PACKING_BOX_NO"]) == Convert.ToString(ISG_LIST.GetCellValue(i, ISG_LIST.GetColumnToIndex("IN_BOX_NO"))))
                        {
                            //이미 존재 합니다.
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10057"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    IDA_LIST.AddUnder();

                    ISG_LIST.SetCellValue("QTY", row["QTY"]);
                    ISG_LIST.SetCellValue("IN_BOX_NO", row["PACKING_BOX_NO"]);
                    ISG_LIST.SetCellValue("CUST_BOX_NO", row["CUST_BARCODE"]);
                }

                for (int i = 0; i < ISG_LIST.RowCount; i++)
                {
                    vTotalQty = vTotalQty + iString.ISDecimaltoZero(ISG_LIST.GetCellValue(i, ISG_LIST.GetColumnToIndex("QTY")));
                }

                V_TOTAL_QTY.EditValue = vTotalQty;

                XLPrinting(ISG_LIST, IDA_TARGET);

                V_BARCODE.EditValue = "";

                V_BARCODE.Focus();
            }
        }

        private void XLPrinting(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {
            bool isError = false;
            string vMessageText = string.Empty;
            string vSaveFileName = string.Empty;
            int vCount = 0;

            int vCountRowDB = pAdapter.OraSelectData.Rows.Count;

            if (vCountRowDB < 1)
            {
                vMessageText = string.Format("Without Data");
                isAppInterfaceAdv1.OnAppMessage(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            System.Windows.Forms.Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            System.Windows.Forms.Application.DoEvents();

            int vPageNumber = 0;
            int vTerritory = GetTerritory(isAppInterfaceAdv1.AppInterface.OraConnectionInfo.TerritoryLanguage);

            vMessageText = string.Format(" Printing Starting");
            isAppInterfaceAdv1.OnAppMessage(vMessageText);
            System.Windows.Forms.Application.DoEvents();

            XLPrinting xlPrinting = new XLPrinting(isAppInterfaceAdv1.AppInterface, isMessageAdapter1);

            try
            {
                vMessageText = string.Empty;
                string vPrintingDate = string.Format("{0:D2}/{1:D2}", System.DateTime.Now.Month, System.DateTime.Now.Day);
                string vPrintingUser = isAppInterfaceAdv1.AppInterface.DisplayName;

                xlPrinting.OpenFileNameExcel = "SOMF0661_LABEL_001.xlsx";

                bool isOpen = xlPrinting.XLFileOpen();

                if (isOpen == true)
                {
                    //for (int i = 0; i < pGrid.RowCount; i++)
                    //{
                        
                    //}
                    vPageNumber = xlPrinting.LineWrite(pAdapter.OraSelectData.Rows[0]);

                    System.Threading.Thread.Sleep(2000);
                    //-------------------------------------------------------------------------------------
                    xlPrinting.Dispose();
                    //-------------------------------------------------------------------------------------
                }
                else
                {
                    vMessageText = "Excel File Open Error";
                }
                //-------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                isError = true;
                vMessageText = ex.Message;
                xlPrinting.Dispose();
            }

            if (isError != true)
            {
                //-------------------------------------------------------------------------------------
                vMessageText = string.Format("{0} Printing End [Total Page : {1}]", vMessageText, vPageNumber);
                isAppInterfaceAdv1.AppInterface.OnAppMessageEvent(vMessageText);
                System.Windows.Forms.Application.DoEvents();
                //-------------------------------------------------------------------------------------
            }
            else
            {
                MessageBoxAdv.Show(vMessageText, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            xlPrinting.KillProcess_Excel();

            System.Windows.Forms.Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            System.Windows.Forms.Application.DoEvents();
        }

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

        #endregion;

        private void isButton1_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_TARGET.SetSelectParamValue("W_OUT_BOX_FLAG", "Y");

            IDA_TARGET.Fill();

            XLPrinting(ISG_LIST, IDA_TARGET);
        }
    }

}
