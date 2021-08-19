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

namespace SOMF0661
{
    public partial class SOMF0661 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public SOMF0661()
        {
            InitializeComponent();
        }

        public SOMF0661(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----


        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    //if (Convert.ToString(iedBILL_TO_CUST_SITE_ID.EditValue) == "")
                    //{
                    //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10055").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;
                    //}
                    IDA_DO_LIST.OraSelectData.AcceptChanges();
                    IDA_DO_LIST.Refillable = true;

                    S_DELIVERY_TYPE_LCODE.EditValue = null;
                    S_ORDER_TYPE_ID.EditValue = null;
                    S_SHIP_METHOD_LCODE.EditValue = null;
                    S_SHIP_TO_CUST_SITE_ID.EditValue = null;

                    IDA_DO_LIST.Fill();

                    iedBILL_TO_CUST_SITE_ID.EditValue = null;
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

        private void SOMF0661_Load(object sender, EventArgs e)
        {
            IDA_DO_LIST.FillSchema();

            iedDELIVERY_DATE_FR.EditValue = DateTime.Today;
            iedDELIVERY_DATE_TO.EditValue = DateTime.Today;
        }

        private void IBT_DELIVERY_ORDER_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int vCheck = 0;

            IDC_TEMP_SEQ.ExecuteNonQuery();

            for (int i = 0; i < ISG_DO_LIST.RowCount; i++)
            {
                if (Convert.ToString(ISG_DO_LIST.GetCellValue(i, ISG_DO_LIST.GetColumnToIndex("SELECT_FLAG"))) == "Y")
                {
                    IDC_TEMP_INSERT.SetCommandParamValue("P_TEMP_SEQ", IDC_TEMP_SEQ.GetCommandParamValue("X_TEMP_SEQ"));
                    IDC_TEMP_INSERT.SetCommandParamValue("P_DELIVERY_ORDER_ID", ISG_DO_LIST.GetCellValue(i, ISG_DO_LIST.GetColumnToIndex("DELIVERY_ORDER_ID")));

                    IDC_TEMP_INSERT.ExecuteNonQuery();
                    
                    vCheck++;
                }
            }

            if (vCheck == 0)
            {
                //선택된 지시가 없습니다.
            }
            else
            {

                Form vSOMF0661_1 = new SOMF0661_1(this.MdiParent, isAppInterfaceAdv1.AppInterface, iedBILL_TO_CUST_SITE_ID.EditValue, S_DELIVERY_TYPE_LCODE.EditValue,
                                                                                                   S_ORDER_TYPE_ID.EditValue, S_SHIP_METHOD_LCODE.EditValue,
                                                                                                   S_SHIP_TO_CUST_SITE_ID.EditValue,
                                                                                                   S_INVOICE_ID.EditValue, IDC_TEMP_SEQ.GetCommandParamValue("X_TEMP_SEQ"));

                vSOMF0661_1.ShowDialog();

                vSOMF0661_1.Dispose();

                IDA_DO_LIST.Refillable = true;
                IDA_DO_LIST.Fill();

                S_DELIVERY_TYPE_LCODE.EditValue = null;
                S_ORDER_TYPE_ID.EditValue = null;
                S_SHIP_METHOD_LCODE.EditValue = null;
                S_SHIP_TO_CUST_SITE_ID.EditValue = null;
                S_INVOICE_ID.EditValue = null;
                iedBILL_TO_CUST_SITE_ID.EditValue = null;
            }

        }

        private void ilaPLIST_SelectedRowData(object pSender)
        {
            if (Convert.ToString(iedBILL_TO_CUST_SITE_ID.EditValue) == "")
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10055").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            IDA_DO_LIST.OraSelectData.AcceptChanges();
            IDA_DO_LIST.Refillable = true;

            S_DELIVERY_TYPE_LCODE.EditValue = null;
            S_ORDER_TYPE_ID.EditValue = null;
            S_SHIP_METHOD_LCODE.EditValue = null;
            S_SHIP_TO_CUST_SITE_ID.EditValue = null;

            Line_Setting();

            iedDELIVERY_ORDER_NO.Focus();
        }

        private void ICB_ALL_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            for (int i = 0; i < ISG_DO_LIST.RowCount; i++)
            {
                ISG_DO_LIST.SetCellValue(i,ISG_DO_LIST.GetColumnToIndex("SELECT_FLAG"), ICB_ALL.CheckBoxValue);
            }
        }

        private void ISG_DO_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            switch (ISG_DO_LIST.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "SELECT_FLAG":

                    if (Convert.ToString(e.NewValue) == "Y")
                    {
                        long vBILL_TO_CUST_SITE_ID = Convert.ToInt64(ISG_DO_LIST.GetCellValue(e.RowIndex, ISG_DO_LIST.GetColumnToIndex("BILL_TO_CUST_SITE_ID")));
                        long vSHIP_TO_CUST_SITE_ID = Convert.ToInt64(ISG_DO_LIST.GetCellValue(e.RowIndex, ISG_DO_LIST.GetColumnToIndex("SHIP_TO_CUST_SITE_ID")));
                        long vORDER_TYPE_ID = Convert.ToInt64(ISG_DO_LIST.GetCellValue(e.RowIndex, ISG_DO_LIST.GetColumnToIndex("ORDER_TYPE_ID")));
                        long vINVOICE_ID = Convert.ToInt64(ISG_DO_LIST.GetCellValue(e.RowIndex, ISG_DO_LIST.GetColumnToIndex("INVOICE_ID")));
                        string vSHIP_METHOD_LCODE = Convert.ToString(ISG_DO_LIST.GetCellValue(e.RowIndex, ISG_DO_LIST.GetColumnToIndex("SHIP_METHOD_LCODE")));
                        //string vDELIVERY_TYPE_LCODE = Convert.ToString(ISG_DO_LIST.GetCellValue(e.RowIndex, ISG_DO_LIST.GetColumnToIndex("DELIVERY_TYPE_LCODE")));

                        decimal test = iString.ISDecimaltoZero(S_SHIP_TO_CUST_SITE_ID.EditValue);

                        if (iString.ISDecimaltoZero(iedBILL_TO_CUST_SITE_ID.EditValue) == 0 &&
                            iString.ISDecimaltoZero(S_SHIP_TO_CUST_SITE_ID.EditValue) == 0 &&
                            iString.ISDecimaltoZero(S_ORDER_TYPE_ID.EditValue) == 0 &&
                            iString.ISDecimaltoZero(S_INVOICE_ID.EditValue) == 0 &&
                            Convert.ToString(S_SHIP_METHOD_LCODE.EditValue) == "")
                        {
                            iedBILL_TO_CUST_SITE_ID.EditValue = vBILL_TO_CUST_SITE_ID;
                            S_SHIP_TO_CUST_SITE_ID.EditValue = vSHIP_TO_CUST_SITE_ID;
                            S_ORDER_TYPE_ID.EditValue = vORDER_TYPE_ID;
                            S_INVOICE_ID.EditValue = vINVOICE_ID;
                            S_SHIP_METHOD_LCODE.EditValue = vSHIP_METHOD_LCODE;
                            //S_DELIVERY_TYPE_LCODE.EditValue = vDELIVERY_TYPE_LCODE;
                        }
                        else
                        {
                            if (vBILL_TO_CUST_SITE_ID == Convert.ToInt64(iedBILL_TO_CUST_SITE_ID.EditValue) &&
                                vSHIP_TO_CUST_SITE_ID == Convert.ToInt64(S_SHIP_TO_CUST_SITE_ID.EditValue) &&
                                vORDER_TYPE_ID == Convert.ToInt64(S_ORDER_TYPE_ID.EditValue) &&
                                vINVOICE_ID == Convert.ToInt64(S_INVOICE_ID.EditValue) &&
                                vSHIP_METHOD_LCODE == Convert.ToString(S_SHIP_METHOD_LCODE.EditValue))
                            {

                            }
                            else
                            {
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10061").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                ISG_DO_LIST.SetCellValue("SELECT_FLAG", "N");
                                return;
                            }
                        }
                    }
                    else
                    {
                        int vChk = 0;

                        for (int i = 0; i < ISG_DO_LIST.RowCount; i++)
                        {
                            string vSelectFalg = Convert.ToString(ISG_DO_LIST.GetCellValue(i, ISG_DO_LIST.GetColumnToIndex("SELECT_FLAG")));
                            if (vSelectFalg == "Y")
                            {
                                vChk++;
                            }
                        }

                        if (vChk == 0)
                        {
                            iedBILL_TO_CUST_SITE_ID.EditValue = null;
                            S_DELIVERY_TYPE_LCODE.EditValue = null;
                            S_ORDER_TYPE_ID.EditValue = null;
                            S_INVOICE_ID.EditValue = null;
                            S_SHIP_METHOD_LCODE.EditValue = null;
                            S_SHIP_TO_CUST_SITE_ID.EditValue = null;
                        }
                    }

                    break;

                default:
                    break;
            }
        }

        private void iedDELIVERY_ORDER_NO_KeyDown(object pSender, KeyEventArgs e)
        {
            string vInputDeliveryOrderNo = iString.ISNull(iedDELIVERY_ORDER_NO.EditValue);

            if (e.KeyCode == Keys.Enter)
            {
                //if (Convert.ToString(iedBILL_TO_CUST_SITE_ID.EditValue) == "")
                //{
                //    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10055").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return;
                //}

                for (int vListRow = 0; vListRow < ISG_DO_LIST.RowCount; vListRow++)
                {
                    string vDeliveryOrderNo = iString.ISNull(ISG_DO_LIST.GetCellValue(vListRow, ISG_DO_LIST.GetColumnToIndex("DELIVERY_ORDER_NO")));

                    if (vInputDeliveryOrderNo == vDeliveryOrderNo)
                    {
                        string X_RESULT_MSG = "[" + vInputDeliveryOrderNo + "] - " + isMessageAdapter1.ReturnText("EAPP_10057");
                        MessageBoxAdv.Show(X_RESULT_MSG, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        iedDELIVERY_ORDER_NO.EditValue = string.Empty;
                        return;
                    }
                }

                IDA_DO_LIST_1.Fill();
                
                int vCount = IDA_DO_LIST_1.SelectRows.Count;

                if (vCount == 0)
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10062").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //해당 지시는 이미 출고되었거나, 존재하지 않습니다.
                }

                foreach (System.Data.DataRow vRow in IDA_DO_LIST_1.OraSelectData.Rows)
                {
                    IDA_DO_LIST.AddUnder();

                    ISG_DO_LIST.SetCellValue("SELECT_FLAG", "Y");
                    ISG_DO_LIST.SetCellValue("DELIVERY_ORDER_NO", vRow["DELIVERY_ORDER_NO"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_DATE", vRow["DELIVERY_DATE"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_REQUEST_DATE", vRow["DELIVERY_REQUEST_DATE"]);
                    ISG_DO_LIST.SetCellValue("ORDER_NO", vRow["ORDER_NO"]);
                    ISG_DO_LIST.SetCellValue("ITEM_CODE", vRow["ITEM_CODE"]);
                    ISG_DO_LIST.SetCellValue("ITEM_DESCRIPTION", vRow["ITEM_DESCRIPTION"]);
                    ISG_DO_LIST.SetCellValue("CUST_SITE_CODE", vRow["CUST_SITE_CODE"]);
                    ISG_DO_LIST.SetCellValue("CUST_SITE_FULL_NAME", vRow["CUST_SITE_FULL_NAME"]);
                    ISG_DO_LIST.SetCellValue("SHIP_TO_CUST_SITE_NAME", vRow["SHIP_TO_CUST_SITE_NAME"]);
                    ISG_DO_LIST.SetCellValue("ORDER_TYPE_DESC", vRow["ORDER_TYPE_DESC"]);
                    ISG_DO_LIST.SetCellValue("SHIP_METHOD_TYPE_DESC", vRow["SHIP_METHOD_TYPE_DESC"]);
                    ISG_DO_LIST.SetCellValue("ITEM_SECTION_DESC", vRow["ITEM_SECTION_DESC"]);
                    ISG_DO_LIST.SetCellValue("UOM_CODE", vRow["UOM_CODE"]);
                    ISG_DO_LIST.SetCellValue("SALES_ORDER_QTY", vRow["SALES_ORDER_QTY"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_REMAIN_QTY", vRow["DELIVERY_REMAIN_QTY"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_TYPE_LCODE", vRow["DELIVERY_TYPE_LCODE"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_TYPE_DESC", vRow["DELIVERY_TYPE_DESC"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_QTY", vRow["DELIVERY_QTY"]);
                    ISG_DO_LIST.SetCellValue("PRE_DELIVERY_QTY", vRow["PRE_DELIVERY_QTY"]);
                    ISG_DO_LIST.SetCellValue("REMAIN_DO_QTY", vRow["REMAIN_DO_QTY"]);
                    ISG_DO_LIST.SetCellValue("CURRENCY_CODE", vRow["CURRENCY_CODE"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_PRICE", vRow["DELIVERY_PRICE"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_AMOUNT", vRow["DELIVERY_AMOUNT"]);
                    ISG_DO_LIST.SetCellValue("CUSTOMER_REV", vRow["CUSTOMER_REV"]);
                    ISG_DO_LIST.SetCellValue("CUST_PO_NO", vRow["CUST_PO_NO"]);
                    ISG_DO_LIST.SetCellValue("CUST_PO_LINE_NO", vRow["CUST_PO_LINE_NO"]);
                    ISG_DO_LIST.SetCellValue("CUST_DO_NO", vRow["CUST_DO_NO"]);
                    ISG_DO_LIST.SetCellValue("CUST_DO_LINE_NO", vRow["CUST_DO_LINE_NO"]);
                    ISG_DO_LIST.SetCellValue("ORDER_HEADER_ID", vRow["ORDER_HEADER_ID"]);
                    ISG_DO_LIST.SetCellValue("ORDER_LINE_ID", vRow["ORDER_LINE_ID"]);
                    ISG_DO_LIST.SetCellValue("ORDER_LINE_NO", vRow["ORDER_LINE_NO"]);
                    ISG_DO_LIST.SetCellValue("REMARK", vRow["REMARK"]);
                    ISG_DO_LIST.SetCellValue("INVENTORY_ITEM_ID", vRow["INVENTORY_ITEM_ID"]);
                    ISG_DO_LIST.SetCellValue("BILL_TO_CUST_SITE_ID", vRow["BILL_TO_CUST_SITE_ID"]);
                    ISG_DO_LIST.SetCellValue("SHIP_TO_CUST_SITE_ID", vRow["SHIP_TO_CUST_SITE_ID"]);
                    ISG_DO_LIST.SetCellValue("ORDER_TYPE_ID", vRow["ORDER_TYPE_ID"]);
                    ISG_DO_LIST.SetCellValue("DELIVERY_ORDER_ID", vRow["DELIVERY_ORDER_ID"]);
                    ISG_DO_LIST.SetCellValue("PO_LINE_ID", vRow["PO_LINE_ID"]);
                    ISG_DO_LIST.SetCellValue("SHIP_METHOD_LCODE", vRow["SHIP_METHOD_LCODE"]);

                    long vBILL_TO_CUST_SITE_ID = Convert.ToInt64(ISG_DO_LIST.GetCellValue(ISG_DO_LIST.GetColumnToIndex("BILL_TO_CUST_SITE_ID")));
                    long vSHIP_TO_CUST_SITE_ID = Convert.ToInt64(ISG_DO_LIST.GetCellValue(ISG_DO_LIST.GetColumnToIndex("SHIP_TO_CUST_SITE_ID")));
                    long vORDER_TYPE_ID = Convert.ToInt64(ISG_DO_LIST.GetCellValue(ISG_DO_LIST.GetColumnToIndex("ORDER_TYPE_ID")));
                    string vSHIP_METHOD_LCODE = Convert.ToString(ISG_DO_LIST.GetCellValue(ISG_DO_LIST.GetColumnToIndex("SHIP_METHOD_LCODE")));
                    string vDELIVERY_TYPE_LCODE = Convert.ToString(ISG_DO_LIST.GetCellValue(ISG_DO_LIST.GetColumnToIndex("DELIVERY_TYPE_LCODE")));

                    if (iString.ISDecimaltoZero(iedBILL_TO_CUST_SITE_ID.EditValue) == 0)
                    {
                        iedBILL_TO_CUST_SITE_ID.EditValue = vBILL_TO_CUST_SITE_ID;
                    }

                    if (iString.ISDecimaltoZero(S_SHIP_TO_CUST_SITE_ID.EditValue) == 0 &&
                        iString.ISDecimaltoZero(S_ORDER_TYPE_ID.EditValue) == 0 &&
                        Convert.ToString(S_SHIP_METHOD_LCODE.EditValue) == "")
                    {
                        S_SHIP_TO_CUST_SITE_ID.EditValue = vSHIP_TO_CUST_SITE_ID;
                        S_ORDER_TYPE_ID.EditValue = vORDER_TYPE_ID;
                        S_SHIP_METHOD_LCODE.EditValue = vSHIP_METHOD_LCODE;
                        S_DELIVERY_TYPE_LCODE.EditValue = vDELIVERY_TYPE_LCODE;
                    }
                    else
                    {
                        if (vBILL_TO_CUST_SITE_ID == Convert.ToInt64(iedBILL_TO_CUST_SITE_ID.EditValue) &&
                            vSHIP_TO_CUST_SITE_ID == Convert.ToInt64(S_SHIP_TO_CUST_SITE_ID.EditValue) &&
                            vORDER_TYPE_ID == Convert.ToInt64(S_ORDER_TYPE_ID.EditValue) &&
                            vSHIP_METHOD_LCODE == Convert.ToString(S_SHIP_METHOD_LCODE.EditValue))
                        {

                        }
                        else
                        {
                            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10061").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ISG_DO_LIST.SetCellValue("SELECT_FLAG", "N");
                            iedDELIVERY_ORDER_NO.EditValue = string.Empty;
                            return;
                        }
                    }
                }
                iedDELIVERY_ORDER_NO.EditValue = string.Empty;
            }
        }

        private void Line_Setting()
        {
            IDA_DO_LIST.SetSelectParamValue("W_SOB_ID", -1);
            IDA_DO_LIST.Fill();

            IDA_DO_LIST.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
        }

        private void ISG_DO_LIST_CellDoubleClick(object pSender)
        {
            Form vSOMF0661_LABEL = new SOMF0661_LABEL(this.MdiParent, isAppInterfaceAdv1.AppInterface, ISG_DO_LIST.GetCellValue("DELIVERY_ORDER_ID"), ISG_DO_LIST.GetCellValue("SHIP_TO_CUST_SITE_NAME"));

            vSOMF0661_LABEL.ShowDialog();

            vSOMF0661_LABEL.Dispose();
        }
    }
}