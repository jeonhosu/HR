using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace PPMF0638
{
    public partial class PPMF0638 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        System.Windows.Forms.PrintDialog PD = new PrintDialog();
        System.Drawing.Printing.PrinterSettings PS = new System.Drawing.Printing.PrinterSettings();

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);

        #region ----- Variables -----

        object oOnhandId;
        object oBoxTypeDesc;
        object oBoxTypeId;
        object oWarehouseName;
        object oLoocationName;
        object oItemCode;
        object oItemDesc;
        object oJobNo;
        object oWeekNum;
        object oSelectFlag;
        object oCustomer;

        string vWeekNum = "";
        string vItemId = "";

        bool vStatus;

        bool bStatus = false;

        #endregion;

        #region ----- Constructor -----

        public PPMF0638()
        {
            InitializeComponent();
        }

        public PPMF0638(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----


        #endregion;

        #region ----- PF4i BarCode Printing Method ----

        private void BarCodePrinting()
        {
            int vRowCount = 0;
            for (int vRow = 0; vRow < ISG_PACKING_LIST.RowCount; vRow++)
            {
                vRowCount++;
            }
            if (vRowCount < 1)
            {
                isAppInterfaceAdv1.OnAppMessage("Print Data is not found. Exit Print");
                return;
            }



            //IDA_PACKING.Fill();
            //int vRowCount = IDA_PACKING.CurrentRows.Count;
            //if (vRowCount < 1)
            //{
            //    isAppInterfaceAdv1.OnAppMessage("Print Data is not found. Exit Print");
            //    return;
            //}

            //라벨 인쇄 위한 개체 선언 //



            //for (int vLineRow = 0; vLineRow < ISG_PACKING_LIST.RowCount; vLineRow++)
            //{
                //라벨 인쇄 위한 개체 선언 //

                PF4i_0110 vBarCodePrint = new PF4i_0110(isAppInterfaceAdv1.AppInterface, printDialog1, printPreviewDialog1);

                foreach (System.Data.DataRow vRow in IDA_PACKING_PRINT.CurrentRows)
                {
                    //string userName = isAppInterfaceAdv1.DISPLAY_NAME.ToString();
                    string vSelectFlag = Convert.ToString(vRow["SELECT_FLAG"]);

                    if (vSelectFlag == "Y")
                    {
                        vBarCodePrint.PRINTING(vRow);
                    }
                }

                try
                {
                    if (vBarCodePrint != null)
                    {
                        vBarCodePrint.Dispose();
                    }
                }
                catch (System.Exception ex)
                {
                    isAppInterfaceAdv1.OnAppMessage(ex.Message);
                }
            //}
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv1.TabIndex)
                    {
                        ISG_NOT_PACKING_LIST.LastConfirmChanges();
                        IDA_NOT_PACKING_LIST.OraSelectData.AcceptChanges();
                        ISG_TEMP_PACKING_LIST.LastConfirmChanges();
                        IDA_TEMP_PACKING_LIST.OraSelectData.AcceptChanges();

                        IDA_NOT_PACKING_LIST.Refillable = true;
                        IDA_TEMP_PACKING_LIST.Refillable = true;

                        IDA_NOT_PACKING_LIST.Fill();
                        IDA_TEMP_PACKING_LIST.Fill();

                        vWeekNum = "";
                        vItemId = "";
                        M_TOTAL_QTY.EditValue = 0;
                        M_SPLIT_QTY.EditValue = 0;
                        M_BOX_QTY.EditValue = 0;
                    }
                    else if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv2.TabIndex)
                    {
                        ISG_PACKING_LIST.LastConfirmChanges();
                        IDA_PACKING_LIST.OraSelectData.AcceptChanges();
                        IDA_PACKING_LIST.Refillable = true;
                        IDA_PACKING_LIST.Fill();
                        //DB_Search();
                    }
                    else if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv3.TabIndex)
                    {
                        ISG_PACKING_PRINT.LastConfirmChanges();
                        IDA_PACKING_PRINT.OraSelectData.AcceptChanges();
                        IDA_PACKING_PRINT.Refillable = true;
                        IDA_PACKING_PRINT.Fill();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv2.TabIndex)
                    {
                        LineValue_Get();

                        if (bStatus == false)
                        {
                            return;
                        }

                        IDA_PACKING_LIST.AddOver();
                        LineValue_Set();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv2.TabIndex)
                    {
                        LineValue_Get();

                        if (bStatus == false)
                        {
                            return;
                        }

                        IDA_PACKING_LIST.AddUnder();
                        LineValue_Set();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv2.TabIndex)
                    {
                        IDA_PACKING_LIST.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv1.TabIndex)
                    {
                        IDA_NOT_PACKING_LIST.Cancel();
                    }
                    else if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv2.TabIndex)
                    {
                        IDA_PACKING_LIST.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv2.TabIndex)
                    {
                        ISG_PACKING_LIST.CurrentCellMoveTo(0, 1);
                        ISG_PACKING_LIST.Focus();
                        ISG_PACKING_LIST.CurrentCellActivate(0, 1);

                        string defaultPrint = GetDefaultPrinter();

                        System.Windows.Forms.DialogResult vResult = printDialog1.ShowDialog();
                        short vInput_Copies = printDialog1.PrinterSettings.DefaultPageSettings.PrinterSettings.Copies;
                        PD.PrinterSettings = printDialog1.PrinterSettings;

                        SetDefaultPrinter(PD.PrinterSettings.PrinterName.ToString());

                        if (Convert.ToString(vResult).Equals("OK"))
                        {
                            XLPrinting(ISG_PACKING_LIST, IDA_PACKING_LIST);
                        }

                        SetDefaultPrinter(defaultPrint);
                    }
                    else if (isTabAdv1.SelectedTab.TabIndex == tabPageAdv3.TabIndex)
                    {
                        ISG_PACKING_PRINT.CurrentCellMoveTo(0, 1);
                        ISG_PACKING_PRINT.Focus();
                        ISG_PACKING_PRINT.CurrentCellActivate(0, 1);

                        string defaultPrint = GetDefaultPrinter();

                        System.Windows.Forms.DialogResult vResult = printDialog1.ShowDialog();
                        short vInput_Copies = printDialog1.PrinterSettings.DefaultPageSettings.PrinterSettings.Copies;
                        PD.PrinterSettings = printDialog1.PrinterSettings;

                        SetDefaultPrinter(PD.PrinterSettings.PrinterName.ToString());

                        if (Convert.ToString(vResult).Equals("OK"))
                        {
                            XLPrinting(ISG_PACKING_PRINT, IDA_PACKING_PRINT);
                        }

                        SetDefaultPrinter(defaultPrint);
                    }
                }
            }
        }

        #endregion;

        public string GetDefaultPrinter()
        {
            PrintDocument PD = new PrintDocument();
            return PD.PrinterSettings.PrinterName;
        }

        private void DB_Search()
        {
            IDA_PACKING_LIST.Refillable = true;
            IDA_PACKING_LIST.Fill();
        }

        private void LineValue_Get()
        {
            oSelectFlag = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("SELECT_FLAG"));
            oWarehouseName = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("WAREHOUSE_NAME"));
            oLoocationName = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("LOCATION_NAME"));
            oItemCode = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("ITEM_CODE"));
            oItemDesc = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("ITEM_DESCRIPTION"));
            oBoxTypeDesc = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("BOX_TYPE_DESC"));
            oJobNo = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("JOB_NO"));
            oWeekNum = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("WEEK_NUM"));
            oCustomer = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("CUSTOMER_DESC"));
            oOnhandId = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("ONHAND_ID"));
            oBoxTypeId = ISG_PACKING_LIST.GetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("BOX_TYPE_ID"));

            if (iString.ISDecimaltoZero(oOnhandId) == 0)
            {
                bStatus = false;
                return;
            }
            else
            {
                bStatus = true;
            }

            if (Convert.ToString(oSelectFlag) == "N")
            {
                bStatus = false;
                return;
            }
            else
            {
                bStatus = true;
            }
        }

        private void LineValue_Set()
        {
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("SELECT_FLAG"), oSelectFlag);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("WAREHOUSE_NAME"), oWarehouseName);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("LOCATION_NAME"), oLoocationName);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("ITEM_CODE"), oItemCode);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("ITEM_DESCRIPTION"), oItemDesc);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("BOX_TYPE_DESC"), oBoxTypeDesc);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("JOB_NO"), oJobNo);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("WEEK_NUM"), oWeekNum);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("CUSTOMER_DESC"), oCustomer);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("X_ONHAND_ID"), oOnhandId);
            ISG_PACKING_LIST.SetCellValue(ISG_PACKING_LIST.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("BOX_TYPE_ID"), oBoxTypeId);
        }

        private void PPMF0638_Load(object sender, EventArgs e)
        {
            IDA_NOT_PACKING_LIST.FillSchema();
            IDA_TEMP_PACKING_LIST.FillSchema();
            IDA_PACKING_LIST.FillSchema();
            IDA_PACKING_PRINT.FillSchema();

            idcFG_WIP.ExecuteNonQuery();
        }

        private void ISG_ONHAND_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            switch (ISG_NOT_PACKING_LIST.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "SELECT_FLAG":

                    if (Convert.ToString(e.NewValue) == "Y")
                    {
                        string vOldItemId = "";

                        if (vItemId == "")
                        {
                            vItemId = Convert.ToString(ISG_NOT_PACKING_LIST.GetCellValue(e.RowIndex, ISG_NOT_PACKING_LIST.GetColumnToIndex("INVENTORY_ITEM_ID")));
                        }
                        else
                        {
                            vOldItemId = Convert.ToString(ISG_NOT_PACKING_LIST.GetCellValue(e.RowIndex, ISG_NOT_PACKING_LIST.GetColumnToIndex("INVENTORY_ITEM_ID")));

                            if (vItemId != vOldItemId)
                            {
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("INV_10133"), "Information", MessageBoxButtons.OK);
                                ISG_NOT_PACKING_LIST.SetCellValue(e.RowIndex, e.ColIndex, "N");
                                return;
                            }
                        }

                        decimal vStoredQty = Convert.ToDecimal(ISG_NOT_PACKING_LIST.GetCellValue(e.RowIndex, ISG_NOT_PACKING_LIST.GetColumnToIndex("ONHAND_QTY")));

                        decimal vTotalQty = iString.ISDecimaltoZero(M_TOTAL_QTY.EditValue);

                        M_TOTAL_QTY.EditValue = vTotalQty + vStoredQty;
                    }
                    else
                    {
                        decimal vStoredQty = Convert.ToDecimal(ISG_NOT_PACKING_LIST.GetCellValue(e.RowIndex, ISG_NOT_PACKING_LIST.GetColumnToIndex("ONHAND_QTY")));

                        decimal vTotalQty = iString.ISDecimaltoZero(M_TOTAL_QTY.EditValue);

                        M_TOTAL_QTY.EditValue = vTotalQty - vStoredQty;
                    }
                    break;
                //분할수량을 입력하면 박스수량은 자동계산
                case "SPLIT_QTY":

                    if (Convert.ToDecimal(e.NewValue) > 0)
                    {
                        decimal vStoredQty = Convert.ToDecimal(ISG_NOT_PACKING_LIST.GetCellValue(ISG_NOT_PACKING_LIST.RowIndex, ISG_NOT_PACKING_LIST.GetColumnToIndex("ONHAND_QTY")));

                        ISG_NOT_PACKING_LIST.SetCellValue(ISG_NOT_PACKING_LIST.RowIndex, ISG_NOT_PACKING_LIST.GetColumnToIndex("BOX_CNT"), Math.Ceiling(vStoredQty / Convert.ToDecimal(e.NewValue)));
                    }
                    break;
                //박스수량을 입력하면 분할수량을 자동계산
                case "BOX_CNT":

                    if (Convert.ToDecimal(e.NewValue) > 0)
                    {
                        decimal vStoredQty = Convert.ToDecimal(ISG_NOT_PACKING_LIST.GetCellValue(ISG_NOT_PACKING_LIST.RowIndex, ISG_NOT_PACKING_LIST.GetColumnToIndex("ONHAND_QTY")));

                        ISG_NOT_PACKING_LIST.SetCellValue(ISG_NOT_PACKING_LIST.RowIndex, ISG_NOT_PACKING_LIST.GetColumnToIndex("SPLIT_QTY"), Math.Ceiling(vStoredQty / Convert.ToDecimal(e.NewValue)));
                    }
                    break;
                default:
                    break;
            }
        }



        private void IBT_DIVISION_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            //분할하시겠습니까?
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("INV_10118"), "Information", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (ISG_TEMP_PACKING_LIST.RowCount > 0)
                {
                    IDA_TEMP_PACKING_LIST.Cancel();
                }

                decimal vTotalQty = iString.ISDecimaltoZero(M_TOTAL_QTY.EditValue);
                decimal vSplitQty = iString.ISDecimaltoZero(M_SPLIT_QTY.EditValue);
                decimal vBoxQty = 0;
                decimal vRemainQty = 0;
                decimal vMergeQty = 0;
                string vInBoxNo = string.Empty;
                string vStatus = string.Empty;

                IDT_ONHAND_QTY_CHECK.BeginTran();

                for (int vRow = 0; vRow < ISG_NOT_PACKING_LIST.RowCount; vRow++)
                {
                    string vSelectFlag = Convert.ToString(ISG_NOT_PACKING_LIST.GetCellValue(vRow, ISG_NOT_PACKING_LIST.GetColumnToIndex("SELECT_FLAG")));

                    if (vSelectFlag == "Y")
                    {
                        int xDeliveryDate = ISG_NOT_PACKING_LIST.GetColumnToIndex("STORED_DATE");
                        int xWeekNum = ISG_NOT_PACKING_LIST.GetColumnToIndex("WEEK_NUM");
                        int xLotNo = ISG_NOT_PACKING_LIST.GetColumnToIndex("JOB_NO");
                        int xDeliveryQty = ISG_NOT_PACKING_LIST.GetColumnToIndex("ONHAND_QTY");
                        int xItemCode = ISG_NOT_PACKING_LIST.GetColumnToIndex("ITEM_CODE");
                        int xOnhandId = ISG_NOT_PACKING_LIST.GetColumnToIndex("ONHAND_ID");


                        DateTime vDeliveryDate = Convert.ToDateTime(ISG_NOT_PACKING_LIST.GetCellValue(vRow, xDeliveryDate));
                        string vWeekNum = Convert.ToString(ISG_NOT_PACKING_LIST.GetCellValue(vRow, xWeekNum));
                        string vLotNo = Convert.ToString(ISG_NOT_PACKING_LIST.GetCellValue(vRow, xLotNo));
                        decimal vOnhandQty = iString.ISDecimaltoZero(ISG_NOT_PACKING_LIST.GetCellValue(vRow, xDeliveryQty));
                        decimal vOnhandId = iString.ISDecimaltoZero(ISG_NOT_PACKING_LIST.GetCellValue(vRow, xOnhandId));
                        string vItemCode = Convert.ToString(ISG_NOT_PACKING_LIST.GetCellValue(vRow, xItemCode));

                        IDC_ONHAND_QTY_CHECK.SetCommandParamValue("P_WAREHOUSE_ID", ISG_NOT_PACKING_LIST.GetCellValue(vRow, ISG_NOT_PACKING_LIST.GetColumnToIndex("WAREHOUSE_ID")));
                        IDC_ONHAND_QTY_CHECK.SetCommandParamValue("P_LOCATION_ID", ISG_NOT_PACKING_LIST.GetCellValue(vRow, ISG_NOT_PACKING_LIST.GetColumnToIndex("LOCATION_ID")));
                        IDC_ONHAND_QTY_CHECK.SetCommandParamValue("P_INVENTORY_ITEM_ID", ISG_NOT_PACKING_LIST.GetCellValue(vRow, ISG_NOT_PACKING_LIST.GetColumnToIndex("INVENTORY_ITEM_ID")));
                        IDC_ONHAND_QTY_CHECK.SetCommandParamValue("P_DELIVERY_QTY", vOnhandQty);

                        IDC_ONHAND_QTY_CHECK.ExecuteNonQuery();

                        if (IDC_ONHAND_QTY_CHECK.ExcuteError)
                        {
                            //현재 재고량이 입고된 수량보다 부족합니다.
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("WIP_10342") + "Onhand Qty : " + IDC_ONHAND_QTY_CHECK.GetCommandParamValue("X_ONHAND_QTY")
                                                , "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            IDT_ONHAND_QTY_CHECK.RollBack();
                            return;
                        }

                        string vMsgCode = Convert.ToString(IDC_ONHAND_QTY_CHECK.GetCommandParamValue("X_MSG_CODE"));

                        if (vMsgCode == "F")
                        {
                            //현재 재고량이 입고된 수량보다 부족합니다.
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("WIP_10342") + "Onhand Qty : " + IDC_ONHAND_QTY_CHECK.GetCommandParamValue("X_ONHAND_QTY")
                                                , "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            IDT_ONHAND_QTY_CHECK.RollBack();
                            return;
                        }

                        vRemainQty = vOnhandQty;

                        while (vRemainQty > 0)
                        {
                            if (vRemainQty < vSplitQty)
                            {
                                if (vMergeQty > 0)
                                {
                                    if (vMergeQty > vRemainQty)
                                    {
                                        vBoxQty = vRemainQty;
                                        vMergeQty = vMergeQty - vRemainQty;
                                        vRemainQty = 0;

                                        vStatus = "S";
                                    }
                                    else
                                    {
                                        vBoxQty = vMergeQty;
                                        vRemainQty = vRemainQty - vMergeQty;
                                        vMergeQty = 0;

                                        vStatus = "S";
                                    }
                                }
                                else
                                {
                                    vBoxQty = vRemainQty;
                                    vMergeQty = vSplitQty - vRemainQty;
                                    vRemainQty = 0;

                                    vStatus = "D";
                                }
                            }
                            else
                            {
                                if (vMergeQty > 0)
                                {
                                    vBoxQty = vMergeQty;
                                    vRemainQty = vRemainQty - vMergeQty;
                                    vMergeQty = 0;

                                    vStatus = "S";
                                }
                                else
                                {
                                    vBoxQty = vSplitQty;
                                    vRemainQty = vRemainQty - vSplitQty;

                                    vStatus = "D";
                                }
                            }

                            if (vStatus == "D")
                            {
                                IDC_GET_BOX_NO.SetCommandParamValue("P_DELIVERY_DATE", DateTime.Today);
                                IDC_GET_BOX_NO.SetCommandParamValue("P_ITEM_DIVISION_CODE", "GOODS");
                                IDC_GET_BOX_NO.ExecuteNonQuery();

                                vInBoxNo = Convert.ToString(IDC_GET_BOX_NO.GetCommandParamValue("X_BOX_NO"));
                            }

                            IDA_TEMP_PACKING_LIST.AddUnder();

                            ISG_TEMP_PACKING_LIST.SetCellValue(ISG_TEMP_PACKING_LIST.RowIndex, ISG_TEMP_PACKING_LIST.GetColumnToIndex("PACKING_BOX_NO"), vInBoxNo);
                            ISG_TEMP_PACKING_LIST.SetCellValue(ISG_TEMP_PACKING_LIST.RowIndex, ISG_TEMP_PACKING_LIST.GetColumnToIndex("ITEM_CODE"), vItemCode);
                            ISG_TEMP_PACKING_LIST.SetCellValue(ISG_TEMP_PACKING_LIST.RowIndex, ISG_TEMP_PACKING_LIST.GetColumnToIndex("PACKING_QTY"), vBoxQty);
                            ISG_TEMP_PACKING_LIST.SetCellValue(ISG_TEMP_PACKING_LIST.RowIndex, ISG_TEMP_PACKING_LIST.GetColumnToIndex("STORED_DATE"), vDeliveryDate);
                            ISG_TEMP_PACKING_LIST.SetCellValue(ISG_TEMP_PACKING_LIST.RowIndex, ISG_TEMP_PACKING_LIST.GetColumnToIndex("WEEK_NUM"), vWeekNum);
                            ISG_TEMP_PACKING_LIST.SetCellValue(ISG_TEMP_PACKING_LIST.RowIndex, ISG_TEMP_PACKING_LIST.GetColumnToIndex("JOB_NO"), vLotNo);
                            ISG_TEMP_PACKING_LIST.SetCellValue(ISG_TEMP_PACKING_LIST.RowIndex, ISG_TEMP_PACKING_LIST.GetColumnToIndex("ONHAND_ID"), vOnhandId);
                        }

                    }

                }

                IDT_ONHAND_QTY_CHECK.Commit();

            }


            ////////////////////////////////////////////////////

        }
        

        private void ISG_TEMP_PACKING_LIST_CurrentCellEditingComplete(object pSender, ISGridAdvExCellEditingEventArgs e)
        {
            switch (ISG_TEMP_PACKING_LIST.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "PACKING_QTY":

                    //IBT_DIVISION_ButtonClick(pSender, e, e.RowIndex);
                    decimal vSumPackingQty = 0;
                    int xDeliveryQty = ISG_NOT_PACKING_LIST.GetColumnToIndex("ONHAND_QTY");
                    decimal vDeliveryQty = iString.ISDecimaltoZero(M_TOTAL_QTY.EditValue);//Convert.ToDecimal(ISG_NOT_PACKING_LIST.GetCellValue(ISG_NOT_PACKING_LIST.RowIndex, xDeliveryQty));

                    for (int i = 0; i < ISG_TEMP_PACKING_LIST.RowCount; i++)
                    {
                        int xPackingQty = ISG_TEMP_PACKING_LIST.GetColumnToIndex("PACKING_QTY");
                        int xBoxNo = ISG_TEMP_PACKING_LIST.GetColumnToIndex("PACKING_BOX_NO");
                        int xOutBoxNo = ISG_TEMP_PACKING_LIST.GetColumnToIndex("OUT_BOX_NO");

                        decimal vPackingQty = Convert.ToDecimal(ISG_TEMP_PACKING_LIST.GetCellValue(i, xPackingQty));
                        string vBoxNo = Convert.ToString(ISG_TEMP_PACKING_LIST.GetCellValue(i, xBoxNo));
                        string vOutBoxNo = Convert.ToString(ISG_TEMP_PACKING_LIST.GetCellValue(i, xOutBoxNo));

                        if (vBoxNo != vOutBoxNo)
                        {
                            vSumPackingQty = vSumPackingQty + vPackingQty;
                        }
                    }

                    V_GAP_QTY.EditValue = vDeliveryQty - vSumPackingQty;

                    break;


                default:
                    break;
            }
        }

        private void BarcodePrint()
        {
            string vPrintName;

            System.Windows.Forms.DialogResult vResult = printDialog1.ShowDialog();
            short vInput_Copies = printDialog1.PrinterSettings.DefaultPageSettings.PrinterSettings.Copies;
            printDialog1.PrinterSettings.Copies = printDialog1.PrinterSettings.DefaultPageSettings.PrinterSettings.Copies;
            vPrintName = printDialog1.PrinterSettings.PrinterName;

            if (Convert.ToString(vResult).Equals("OK"))
            {
                BarCodePrinting();               
            }   
        }

        private void IBT_DIVISION_CON_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (iString.ISDecimaltoZero(V_GAP_QTY.EditValue) != 0)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("INV_10132"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            vStatus = true;

            try
            {
                IDA_NOT_PACKING_LIST.Update();

                for (int i = 0; i < ISG_TEMP_PACKING_LIST.RowCount; i++)
                {
                    IDC_TEMP_PACKING.SetCommandParamValue("P_ONHAND_ID", ISG_TEMP_PACKING_LIST.GetCellValue(i, ISG_TEMP_PACKING_LIST.GetColumnToIndex("ONHAND_ID")));
                    IDC_TEMP_PACKING.SetCommandParamValue("P_PACKING_BOX_NO", ISG_TEMP_PACKING_LIST.GetCellValue(i, ISG_TEMP_PACKING_LIST.GetColumnToIndex("PACKING_BOX_NO")));
                    IDC_TEMP_PACKING.SetCommandParamValue("P_PACKING_QTY", ISG_TEMP_PACKING_LIST.GetCellValue(i, ISG_TEMP_PACKING_LIST.GetColumnToIndex("PACKING_QTY")));
                    IDC_TEMP_PACKING.SetCommandParamValue("P_BOX_TYPE_ID", ISG_TEMP_PACKING_LIST.GetCellValue(i, ISG_TEMP_PACKING_LIST.GetColumnToIndex("BOX_TYPE_ID")));

                    IDC_TEMP_PACKING.ExecuteNonQuery();
                }

                ISG_NOT_PACKING_LIST.LastConfirmChanges();
                IDA_NOT_PACKING_LIST.OraSelectData.AcceptChanges();

                IDA_NOT_PACKING_LIST.Refillable = true;
                IDA_NOT_PACKING_LIST.Fill();

                vWeekNum = "";
                vItemId = "";
                M_TOTAL_QTY.EditValue = 0;
                M_SPLIT_QTY.EditValue = 0;
                M_BOX_QTY.EditValue = 0;
                
            }
            catch
            {
                vStatus = false;
            }

            if (vStatus == true)
            {
                IDA_PACKING_LIST.SetSelectParamValue("W_ONHAND_STATUS", "ENTER");
                IDA_PACKING_LIST.Fill();
                IDA_PACKING_PRINT.SetSelectParamValue("W_ONHAND_STATUS", "ENTER");
                IDA_PACKING_PRINT.Fill();

                IDC_TEMP_UPDATE.ExecuteNonQuery();

                IDA_PACKING_LIST.SetSelectParamValue("W_ONHAND_STATUS", "");
                IDA_PACKING_PRINT.SetSelectParamValue("W_ONHAND_STATUS", "");

                isTabAdv1.SelectedIndex = 2;
                isTabAdv1.SelectedTab.Focus();
            }
        }

        private void ISG_ONHAND_LIST_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            switch (ISG_PACKING_LIST.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "SELECT_FLAG":

                    decimal vTotalQty = 0;

                    for (int i = 0; i < ISG_NOT_PACKING_LIST.RowCount; i++)
                    {
                        if (Convert.ToString(ISG_NOT_PACKING_LIST.GetCellValue(i, ISG_NOT_PACKING_LIST.GetColumnToIndex("SELECT_FLAG"))) == "Y")
                        {
                            vTotalQty = vTotalQty + iString.ISDecimaltoZero(ISG_NOT_PACKING_LIST.GetCellValue(i, ISG_NOT_PACKING_LIST.GetColumnToIndex("ONHAND_QTY")));
                        }
                    }

                    M_TOTAL_QTY.EditValue = vTotalQty;

                    break;


                case "ONHAND_QTY":

                    int iOnhandId = Convert.ToInt32(ISG_PACKING_LIST.GetCellValue(e.RowIndex, ISG_PACKING_LIST.GetColumnToIndex("X_ONHAND_ID")));

                    for (int i = 0; i < ISG_PACKING_LIST.RowCount; i++)
                    {
                        int iOnhandId_1 = Convert.ToInt32(ISG_PACKING_LIST.GetCellValue(i, ISG_PACKING_LIST.GetColumnToIndex("ONHAND_ID")));
                        int iRowIndex = e.RowIndex;

                        if (iOnhandId == iOnhandId_1 && i != iRowIndex)
                        {
                            decimal dOnhandQty = iString.ISDecimaltoZero(ISG_PACKING_LIST.GetCellValue(i, ISG_PACKING_LIST.GetColumnToIndex("ONHAND_QTY")));
                            decimal nOnhandQty = iString.ISDecimaltoZero(e.CellValue);

                            if (dOnhandQty >= nOnhandQty)
                            {
                                ISG_PACKING_LIST.SetCellValue(i, ISG_PACKING_LIST.GetColumnToIndex("ONHAND_QTY"), (dOnhandQty - nOnhandQty));
                            }
                            else
                            {

                            }
                        }
                    }

                    break;

                default:
                    break;
            }
        }

        private void IBT_MERGE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int vCnt;
            int vCount = 0;
            string vSelectFlag;
            string vOutBoxNo;

            IDC_GET_OUT_BOX_NO.SetCommandParamValue("P_ITEM_DIVISION_CODE","GOODS");

            IDC_GET_OUT_BOX_NO.ExecuteNonQuery();

            vOutBoxNo = Convert.ToString(IDC_GET_OUT_BOX_NO.GetCommandParamValue("X_OUT_BOX_NO"));

            IDT_OUT_BOX_UPDATE.BeginTran();

            for (vCnt = 0; ISG_PACKING_LIST.RowCount > vCnt; vCnt++)
            {
                vSelectFlag = Convert.ToString(ISG_PACKING_LIST.GetCellValue(vCnt, ISG_PACKING_LIST.GetColumnToIndex("SELECT_FLAG")));

                if (vSelectFlag == "Y")
                {
                    ISG_PACKING_LIST.SetCellValue(vCnt, ISG_PACKING_LIST.GetColumnToIndex("OUT_BOX_NO"), vOutBoxNo);

                    IDC_OUT_BOX_UPDATE.SetCommandParamValue("P_ONHAND_ID", ISG_PACKING_LIST.GetCellValue(vCnt, ISG_PACKING_LIST.GetColumnToIndex("ONHAND_ID")));
                    IDC_OUT_BOX_UPDATE.SetCommandParamValue("P_OUT_BOX_NO", ISG_PACKING_LIST.GetCellValue(vCnt, ISG_PACKING_LIST.GetColumnToIndex("OUT_BOX_NO")));

                    IDC_OUT_BOX_UPDATE.ExecuteNonQuery();

                    if (IDC_OUT_BOX_UPDATE.ExcuteError)
                    {
                        IDT_OUT_BOX_UPDATE.RollBack();
                        vCount = 0;
                        return;
                    }

                    vCount++;
                }
            }

            IDT_OUT_BOX_UPDATE.Commit();

            if (vCount > 0)
            {
                DB_Search();

                //for()
            }
        }

        private void IBT_OFF_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (MessageBoxAdv.Show(isMessageAdapter1.ReturnText("INV_10130"), "Information", MessageBoxButtons.YesNo) == DialogResult.Yes)//분할 해제 하시겠습니까?
            {
                IDT_BOX_OFF.BeginTran();

                for(int i = 0; i < ISG_PACKING_LIST.RowCount; i++)
                {
                    IDC_BOX_OFF.SetCommandParamValue("P_SELECT_FLAG", ISG_PACKING_LIST.GetCellValue(i, ISG_PACKING_LIST.GetColumnToIndex("SELECT_FLAG")));
                    IDC_BOX_OFF.SetCommandParamValue("P_ONHAND_ID", ISG_PACKING_LIST.GetCellValue(i, ISG_PACKING_LIST.GetColumnToIndex("ONHAND_ID")));

                    IDC_BOX_OFF.ExecuteNonQuery();

                    if (IDC_BOX_OFF.ExcuteError)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("INV_10129"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        IDT_BOX_OFF.RollBack();
                        return;
                    }
                }

                IDT_BOX_OFF.Commit();

                DB_Search();
            }
        }

        private void ICB_SELECT_ALL_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            for (int i = 0; i < ISG_PACKING_LIST.RowCount; i++)
            {
                string vCustomerDesc = Convert.ToString(ISG_PACKING_LIST.GetCellValue(i, ISG_PACKING_LIST.GetColumnToIndex("CUSTOMER_DESC")));

                if (vCustomerDesc != "")
                {
                    ISG_PACKING_LIST.SetCellValue(i, ISG_PACKING_LIST.GetColumnToIndex("SELECT_FLAG"), ICB_SELECT_ALL.CheckBoxString);
                }
            }
        }

        private void ICB_PRINT_SELECT_ALL_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            for (int i = 0; i < ISG_PACKING_PRINT.RowCount; i++)
            {
                string vCustomerDesc = Convert.ToString(ISG_PACKING_PRINT.GetCellValue(i, ISG_PACKING_PRINT.GetColumnToIndex("CUSTOMER_DESC")));

                if (vCustomerDesc != "")
                {
                    ISG_PACKING_PRINT.SetCellValue(i, ISG_PACKING_PRINT.GetColumnToIndex("SELECT_FLAG"), ICB_PRINT_SELECT_ALL.CheckBoxString);
                }
            }
        }

        private void M_SPLIT_QTY_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            if (Convert.ToDecimal(M_SPLIT_QTY.EditValue) > 0 && Convert.ToDecimal(M_TOTAL_QTY.EditValue) > 0)
            {
                M_BOX_QTY.EditValue = Math.Ceiling(iString.ISDecimaltoZero(M_TOTAL_QTY.EditValue) / iString.ISDecimaltoZero(M_SPLIT_QTY.EditValue));
            }
        }

        private void M_BOX_QTY_CurrentEditValidated(object pSender, ISEditAdvValidatedEventArgs e)
        {
            //if (Convert.ToDecimal(M_BOX_QTY.EditValue) > 0 && Convert.ToDecimal(M_TOTAL_QTY.EditValue) > 0)
            //{
            //    M_SPLIT_QTY.EditValue = Math.Ceiling(iString.ISDecimaltoZero(M_TOTAL_QTY.EditValue) / iString.ISDecimaltoZero(M_BOX_QTY.EditValue));
            //}
        }

        private void ISG_NOT_PACKING_LIST_Click(object sender, EventArgs e)
        {

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

                xlPrinting.OpenFileNameExcel = "PPMF0638_001.xlsx";

                bool isOpen = xlPrinting.XLFileOpen();

                if (isOpen == true)
                {
                    for (int i = 0; i < pGrid.RowCount; i++)
                    {
                        string vSelectFlag = Convert.ToString(pGrid.GetCellValue(i, pGrid.GetColumnToIndex("SELECT_FLAG")));

                        if (vSelectFlag == "Y")
                        {
                            vPageNumber = xlPrinting.LineWrite(pAdapter.OraSelectData.Rows[i]);
                            //int vPageStart = vPageNumber;
                            //xlPrinting.Printing(1, 1);
                        }
                    }

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

        private void ISG_PACKING_PRINT_Click(object sender, EventArgs e)
        {

        }
    }
}