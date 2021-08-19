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
    public partial class SOMF0661_1 : Office2007Form
    {
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        object oDELIVERY_TYPE_LCODE;
        object oORDER_TYPE_ID;
        object oSHIP_METHOD_LCODE;
        object oSHIP_TO_CUST_SITE_ID;
        object oINVOICE_ID;

        #region ----- Variables -----


        #endregion;
        
        #region ----- Constructor -----

        public SOMF0661_1(Form pMainForm, ISAppInterface pAppInterface, object pBILL_TO_CUST_SITE_ID, object pDELIVERY_TYPE_LCODE,
                                                                        object pORDER_TYPE_ID, object pSHIP_METHOD_LCODE,
                                                                        object pSHIP_TO_CUST_SITE_ID, object pINVOICE_ID,
                                                                        object pTEMP_SEQ)
        {
            InitializeComponent();

            isAppInterfaceAdv1.AppInterface = pAppInterface;

            iedBILL_TO_CUST_SITE_ID.EditValue = pBILL_TO_CUST_SITE_ID;
            oSHIP_TO_CUST_SITE_ID = pSHIP_TO_CUST_SITE_ID;
            oORDER_TYPE_ID = pORDER_TYPE_ID;
            oINVOICE_ID = pINVOICE_ID;
            oSHIP_METHOD_LCODE = pSHIP_METHOD_LCODE;
            iedTEMP_SEQ.EditValue = pTEMP_SEQ;
        }

        #endregion;

        #region ----- Events -----

        private void SOMF0661_1_Load(object sender, EventArgs e)
        {
            idaHEADER.FillSchema();
            IDA_DELIVERY_ORDER.FillSchema();
            IDA_BOX_LIST.FillSchema();
            IDA_BOX_NO_WORK_OUT.FillSchema();

            Header_Setting();

            IDA_DELIVERY_ORDER.Fill();

            IBT_TARGET_SERACH_ButtonClick(sender, e);

            iedBOX_NO.Focus();
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

        private void Header_Setting()
        {
            idaHEADER.AddUnder();

            iedPICK_DATE.EditValue = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            iedPICK_PERSON_ID.EditValue = isAppInterfaceAdv1.PERSON_ID;
            iedDISPLAY_NAME.EditValue = isAppInterfaceAdv1.DISPLAY_NAME;
            //iedSHIPMENT_TYPE_ID.EditValue = "SALES_PICK";
            iedCUST_SITE_SHORT_NAME.Focus();

            idcFG_CUST.ExecuteNonQuery();

            idcFG_MAIN.SetCommandParamValue("P_TEMP_SEQ", iedTEMP_SEQ.EditValue);
            idcFG_MAIN.ExecuteNonQuery();
            //iedFROM_WAREHOUSE_ID.EditValue = idcFG_MAIN.GetCommandParamValue("X_WAREHOUSE_ID");
            //iedFROM_WAREHOUSE_CODE.EditValue = idcFG_MAIN.GetCommandParamValue("X_WAREHOUSE_CODE");
            //iedFROM_WAREHOUSE_NAME.EditValue = idcFG_MAIN.GetCommandParamValue("X_WAREHOUSE_NAME");
            //iedFROM_LOCATION_ID.EditValue = idcFG_MAIN.GetCommandParamValue("X_LOCATION_ID");
            //iedFROM_LOCATION_CODE.EditValue = idcFG_MAIN.GetCommandParamValue("X_LOCATION_CODE");
            //iedFROM_LOCATION_NAME.EditValue = idcFG_MAIN.GetCommandParamValue("X_LOCATION_NAME");

            idcFG_SHIP.SetCommandParamValue("P_BILL_TO_CUST_SITE_ID", iedBILL_TO_CUST_SITE_ID.EditValue);
            idcFG_SHIP.ExecuteNonQuery();

            idcINVOICE_NO.SetCommandParamValue("P_INVOICE_ID", oINVOICE_ID);
            idcINVOICE_NO.ExecuteNonQuery();
            //iedTO_WAREHOUSE_ID.EditValue = idcFG_SHIP.GetCommandParamValue("X_WAREHOUSE_ID");
            //iedTO_WAREHOUSE_CODE.EditValue = idcFG_SHIP.GetCommandParamValue("X_WAREHOUSE_CODE");
            //iedTO_WAREHOUSE_NAME.EditValue = idcFG_SHIP.GetCommandParamValue("X_WAREHOUSE_NAME");
            //iedTO_LOCATION_ID.EditValue = idcFG_SHIP.GetCommandParamValue("X_LOCATION_ID");
            //iedTO_LOCATION_CODE.EditValue = idcFG_SHIP.GetCommandParamValue("X_LOCATION_CODE");
            //iedTO_LOCATION_NAME.EditValue = idcFG_SHIP.GetCommandParamValue("X_LOCATION_NAME");
            iedSHIP_TO_CUST_SITE_ID.EditValue = oSHIP_TO_CUST_SITE_ID;
            iedSHIPMENT_TYPE_CODE.EditValue = oORDER_TYPE_ID;
            iedSHIPPING_METHOD_LCODE.EditValue = oSHIP_METHOD_LCODE;

            idcFG_SHIP_TO.ExecuteNonQuery();
            
            idcSHIP_TYPE.ExecuteNonQuery();
            
            idcSHIPPING_METHOD.ExecuteNonQuery();

        }

        private void iedBOX_NO_KeyDown(object pSender, KeyEventArgs e)
        {
            decimal vDeliveryQty = 0;

            if (e.KeyData == Keys.Enter)
            {
                for (int j = 0; j < ISG_BOX_LIST.RowCount; j++)
                {
                    string vBoxNo = Convert.ToString(ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO")));
                    string vOutBoxNo = Convert.ToString(ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("OUT_BOX_NO")));
                    string vSelectFlag = Convert.ToString(ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG")));

                    if (Convert.ToString(iedBOX_NO.EditValue) == vBoxNo || Convert.ToString(iedBOX_NO.EditValue) == vOutBoxNo)
                    {
                        IDC_FIFO_CHECK.SetCommandParamValue("W_INVENTORY_ITEM_ID", ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("INVENTORY_ITEM_ID")));
                        IDC_FIFO_CHECK.SetCommandParamValue("W_BOX_NO", ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO")));

                        IDC_FIFO_CHECK.ExecuteNonQuery();

                        string vErrMsg = Convert.ToString(IDC_FIFO_CHECK.GetCommandParamValue("X_ERR_MSG"));

                        if (vErrMsg != "OK")
                        {
                            MessageBoxAdv.Show(vErrMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        for (int i = 0; i < ISG_ORDER_LIST.RowCount; i++)
                        {
                            string vItemCode = Convert.ToString(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("ITEM_CODE")));

                            if (vItemCode == Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("ITEM_CODE"))))
                            {
                                ISG_BOX_LIST.CurrentCellMoveTo(j, 0);
                                ISG_BOX_LIST.Focus();
                                ISG_BOX_LIST.CurrentCellActivate(j, 0);

                                decimal vOrderRemainDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("REAMAIN_DELIVERY_QTY")));
                                decimal vOrderDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY")));
                                decimal vPackingQty = Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("ONHAND_QTY")));

                                if (Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO"))) == "NOT_PACKING" && vSelectFlag == "N")
                                {
                                    if (vOrderRemainDeliveryQty <= vPackingQty)
                                    {
                                        ISG_BOX_LIST.SetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), vOrderRemainDeliveryQty - vOrderDeliveryQty);
                                    }
                                    else
                                    {
                                        ISG_BOX_LIST.SetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), vPackingQty);
                                    }
                                }
                                else if (Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO"))) != "NOT_PACKING" && vSelectFlag == "N")
                                {
                                    ISG_BOX_LIST.SetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("ONHAND_QTY")));
                                }
                                else if (vSelectFlag == "Y")
                                {
                                    //이미 존재 합니다.
                                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10057").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }

                                vDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) + Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY")));

                                if (vOrderRemainDeliveryQty < vDeliveryQty)
                                {
                                    if (Convert.ToString(ISG_BOX_LIST.GetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO"))) == "NOT_PACKING")
                                    {
                                        ISG_BOX_LIST.SetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), vOrderRemainDeliveryQty - vOrderDeliveryQty);
                                        ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), vOrderRemainDeliveryQty);
                                    }
                                    else
                                    {
                                        //출고수량이 지시잔량 보다 많습니다.
                                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10065").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        ISG_BOX_LIST.SetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG"), "N");
                                        ISG_BOX_LIST.SetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), 0);
                                        return;
                                    }
                                }
                                else
                                {
                                    ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), Convert.ToDecimal(vDeliveryQty));
                                }
                                ISG_BOX_LIST.SetCellValue(j, ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG"), "Y");
                            }
                        }
                    }
                }
                iedBOX_NO.EditValue = "";
                iedBOX_NO.Focus();
            }

            //decimal vSumQty = 0;

            //if (e.KeyData == Keys.Enter)
            //{
            //    IDA_BOX_NO_CHECK.SetSelectParamValue("W_BOX_NO", iedBOX_NO.EditValue);

            //    IDA_BOX_NO_CHECK.Fill();

            //    foreach (DataRow row in IDA_BOX_NO_CHECK.SelectRows)
            //    {
            //        idcBOX_NO_TARGET.SetCommandParamValue("W_PACKING_BOX_NO", row["PACKING_BOX_NO"]);

            //        idcBOX_NO_TARGET.ExecuteNonQuery();

            //        vSumQty = vSumQty + iString.ISDecimaltoZero(idcBOX_NO_TARGET.GetCommandParamValue("X_ONHAND_QTY"));

            //        decimal vDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(ISG_ORDER_LIST.RowIndex, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY")));
            //        decimal vRemainDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(ISG_ORDER_LIST.RowIndex, ISG_ORDER_LIST.GetColumnToIndex("REAMAIN_DELIVERY_QTY")));

            //        if (vRemainDeliveryQty < (vDeliveryQty + vSumQty))
            //        {
            //            MessageBoxAdv.Show("지시잔량보다 출고수량이 많습니다.");
            //            return;
            //        }
            //    }

            //    foreach (DataRow row in IDA_BOX_NO_CHECK.SelectRows)
            //    {
            //        idcBOX_NO_TARGET.SetCommandParamValue("W_PACKING_BOX_NO", row["PACKING_BOX_NO"]);

            //        idcBOX_NO_TARGET.ExecuteNonQuery();

            //        decimal vOnhandQty = Convert.ToDecimal(idcBOX_NO_TARGET.GetCommandParamValue("X_ONHAND_QTY"));

            //        if (vOnhandQty > 0)
            //        {
            //            for (int i = 0; i < ISG_BOX_LIST.RowCount; i++)
            //            {
            //                string vOldPackingBoxNo = Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO")));
            //                string vSelectFlag = Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG")));

            //                if (vOldPackingBoxNo == Convert.ToString(iedBOX_NO.EditValue))
            //                {
            //                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10029", "&&DATA:=Packing Box No").ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //                    return;
            //                }
            //            }

            //            for (int j = 0; j < ISG_ORDER_LIST.RowCount; j++)
            //            {
            //                string vItemCode = Convert.ToString(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("ITEM_CODE")));
            //                if (vItemCode == Convert.ToString(idcBOX_NO_TARGET.GetCommandParamValue("X_ITEM_CODE")))
            //                {
            //                    decimal vDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) + vOnhandQty;
            //                    decimal vRemainDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("REAMAIN_DELIVERY_QTY")));

            //                    if (vRemainDeliveryQty < vDeliveryQty)
            //                    {
            //                        MessageBoxAdv.Show("지시잔량보다 출고수량이 많습니다.");
            //                        return;
            //                    }

            //                    ISG_ORDER_LIST.SetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), vDeliveryQty);
            //                }
            //            }

            //            IDA_BOX_LIST.AddUnder();

            //            ISG_BOX_LIST.SetCellValue("SELECT_FLAG", "Y");
            //            ISG_BOX_LIST.SetCellValue("OUT_BOX_NO", idcBOX_NO_TARGET.GetCommandParamValue("X_OUT_BOX_NO"));
            //            ISG_BOX_LIST.SetCellValue("PACKING_BOX_NO", idcBOX_NO_TARGET.GetCommandParamValue("X_PACKING_BOX_NO"));
            //            ISG_BOX_LIST.SetCellValue("ITEM_CODE", idcBOX_NO_TARGET.GetCommandParamValue("X_ITEM_CODE"));
            //            ISG_BOX_LIST.SetCellValue("ITEM_DESCRIPTION", idcBOX_NO_TARGET.GetCommandParamValue("X_ITEM_DESCRIPTION"));
            //            ISG_BOX_LIST.SetCellValue("ITEM_SPECIFICATION", idcBOX_NO_TARGET.GetCommandParamValue("X_ITEM_SPECIFICATION"));
            //            ISG_BOX_LIST.SetCellValue("ONHAND_QTY", idcBOX_NO_TARGET.GetCommandParamValue("X_ONHAND_QTY"));
            //            ISG_BOX_LIST.SetCellValue("DELIVERY_QTY", idcBOX_NO_TARGET.GetCommandParamValue("X_ONHAND_QTY"));
            //            ISG_BOX_LIST.SetCellValue("INVENTORY_ITEM_ID", idcBOX_NO_TARGET.GetCommandParamValue("X_INVENTORY_ITEM_ID"));
            //            ISG_BOX_LIST.SetCellValue("STORED_TRX_ID", idcBOX_NO_TARGET.GetCommandParamValue("X_STORED_TRX_ID"));
            //            ISG_BOX_LIST.SetCellValue("WIP_JOB_ID", idcBOX_NO_TARGET.GetCommandParamValue("X_WIP_JOB_ID"));

            //            iedBOX_NO.EditValue = "";
            //            iedBOX_NO.Focus();

            //            IBT_TARGET_SEARCH.Enabled = false;
            //        }
            //    }
            //}
            //IDA_DELIVERY_ORDER.OraSelectData.AcceptChanges();
            //IDA_DELIVERY_ORDER.Refillable = true;

            ////IDA_BOX_LIST.OraSelectData.AcceptChanges();
            ////IDA_BOX_LIST.Refillable = true;
        }

        private void IBT_TARGET_SERACH_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_DELIVERY_ORDER.OraSelectData.AcceptChanges();
            IDA_DELIVERY_ORDER.Refillable = true;

            IDA_DELIVERY_ORDER.Fill();

            IDA_BOX_LIST.Refillable = true;

            IDA_BOX_LIST.SetSelectParamValue("W_SOB_ID", -1);
            IDA_BOX_LIST.Fill();

            IDA_BOX_LIST.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);

            for(int k = 0; k < ISG_ORDER_LIST.RowCount; k++)
            {
                object test = ISG_ORDER_LIST.GetCellValue(k, ISG_ORDER_LIST.GetColumnToIndex("INVENTORY_ITEM_ID"));
                IDA_BOX_NO_WORK_OUT.SetSelectParamValue("W_INVENTORY_ITEM_ID", ISG_ORDER_LIST.GetCellValue(k, ISG_ORDER_LIST.GetColumnToIndex("INVENTORY_ITEM_ID")));

                IDA_BOX_NO_WORK_OUT.Fill();

                foreach (DataRow row in IDA_BOX_NO_WORK_OUT.SelectRows)
                {
                    //idcBOX_NO_TARGET.SetCommandParamValue("W_PACKING_BOX_NO", row["PACKING_BOX_NO"]);

                    //idcBOX_NO_TARGET.ExecuteNonQuery();

                    //decimal vOnhandQty = Convert.ToDecimal(idcBOX_NO_TARGET.GetCommandParamValue("X_ONHAND_QTY"));
                    decimal vOnhandQty = Convert.ToDecimal(row["ONHAND_QTY"]);

                    if (vOnhandQty > 0)
                    {
                        //for (int i = 0; i < ISG_BOX_LIST.RowCount; i++)
                        //{
                        //    string vOldPackingBoxNo = Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO")));
                        //    if (vOldPackingBoxNo == Convert.ToString(iedBOX_NO.EditValue))
                        //    {
                        //        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10029", "&&DATA:=Packing Box No").ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //        return;
                        //    }
                        //}

                        IDA_BOX_LIST.AddUnder();

                        ISG_BOX_LIST.SetCellValue("SELECT_FLAG", "N");
                        ISG_BOX_LIST.SetCellValue("WAREHOUSE_NAME", row["WAREHOUSE_NAME"]);
                        ISG_BOX_LIST.SetCellValue("LOCATION_NAME", row["LOCATION_NAME"]);
                        ISG_BOX_LIST.SetCellValue("OUT_BOX_NO", row["OUT_BOX_NO"]);
                        ISG_BOX_LIST.SetCellValue("PACKING_BOX_NO", row["PACKING_BOX_NO"]);
                        ISG_BOX_LIST.SetCellValue("ITEM_CODE", row["ITEM_CODE"]);
                        ISG_BOX_LIST.SetCellValue("ITEM_DESCRIPTION", row["ITEM_DESCRIPTION"]);
                        ISG_BOX_LIST.SetCellValue("ITEM_SPECIFICATION", row["ITEM_SPECIFICATION"]);
                        ISG_BOX_LIST.SetCellValue("ONHAND_QTY", row["ONHAND_QTY"]);
                        ISG_BOX_LIST.SetCellValue("DELIVERY_QTY", 0);
                        ISG_BOX_LIST.SetCellValue("INVENTORY_ITEM_ID", row["INVENTORY_ITEM_ID"]);
                        ISG_BOX_LIST.SetCellValue("STORED_TRX_ID", row["STORED_TRX_ID"]);
                        ISG_BOX_LIST.SetCellValue("WIP_JOB_ID", row["WIP_JOB_ID"]);
                        ISG_BOX_LIST.SetCellValue("LOCATION_ID", row["LOCATION_ID"]);
                        ISG_BOX_LIST.SetCellValue("WAREHOUSE_ID", row["WAREHOUSE_ID"]);
                        ISG_BOX_LIST.SetCellValue("UOM_CODE", row["UOM_CODE"]);
                        ISG_BOX_LIST.SetCellValue("BOM_ITEM_ID", row["BOM_ITEM_ID"]);
                        ISG_BOX_LIST.SetCellValue("ORDER_HEADER_ID", row["ORDER_HEADER_ID"]);
                        ISG_BOX_LIST.SetCellValue("ORDER_LINE_ID", row["ORDER_LINE_ID"]);
                        ISG_BOX_LIST.SetCellValue("WEEK_NUM", row["WEEK_NUM"]);
                        ISG_BOX_LIST.SetCellValue("WEEK_DATE", row["WEEK_DATE"]);

                        iedBOX_NO.EditValue = "";
                        iedBOX_NO.Focus();
                    }

                }

            }

            Set_GRID_STATUS_ROW(ISG_BOX_LIST.GetCellValue("PACKING_BOX_NO"));
        }

        private void IBT_ITEM_ISSUE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(0, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) != 0)
            {
                idaHEADER.Update();

                IBT_ITEM_ISSUE.Enabled = false;
                IBT_TARGET_SEARCH.Enabled = false;
                iedBOX_NO.Enabled = false;
                ICB_ALL.Enabled = false;
                ISG_BOX_LIST.GridAdvExColElement[ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG")].Insertable = 0;
                ISG_BOX_LIST.GridAdvExColElement[ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG")].Updatable = 0;
                ISG_BOX_LIST.GridAdvExColElement[ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY")].Insertable = 0;
                ISG_BOX_LIST.GridAdvExColElement[ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY")].Updatable = 0;
            }
        }

        private void ilaFROM_WAREHOUSE_SelectedRowData(object pSender)
        {
            IDA_DELIVERY_ORDER.OraSelectData.AcceptChanges();
            IDA_DELIVERY_ORDER.Refillable = true;

            IDA_DELIVERY_ORDER.Fill();

            IDA_BOX_LIST.Refillable = true;

            IDA_BOX_LIST.SetSelectParamValue("W_SOB_ID", -1);
            IDA_BOX_LIST.Fill();

            IDA_BOX_LIST.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);

            IBT_TARGET_SEARCH.Enabled = true;
        }
        //private void C_SELECT_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        //{
        //    for (int vLoop = 0; vLoop < ISG_PURCHASE_LINE.RowCount; vLoop++)
        //    {
        //        ISG_PURCHASE_LINE.SetCellValue(vLoop, 0, C_SELECT_FLAG.CheckBoxValue.ToString());
        //        //IDA_DISPOSAL_LINE.OraSelectData.AcceptChanges();
        //        //IDA_DISPOSAL_LINE.Refillable = true;
        //        if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
        //        {
        //            ISGridAdvExChangedEventArgs vISGridAdvExChangedEventArgs = new ISGridAdvExChangedEventArgs(vLoop, 0, "N", "Y");
        //            ISG_DISPOSAL_LINE_CurrentCellChanged(this, vISGridAdvExChangedEventArgs);
        //        }
        //        else
        //        {
        //            ISGridAdvExChangedEventArgs vISGridAdvExChangedEventArgs = new ISGridAdvExChangedEventArgs(vLoop, 0, "Y", "N");
        //            ISG_DISPOSAL_LINE_CurrentCellChanged(this, vISGridAdvExChangedEventArgs);
        //        }
        //    }
        //}


        private void ICB_ALL_CheckedChange(object pSender, ISCheckEventArgs e)
        {

            for (int vLoop = 0; vLoop < ISG_BOX_LIST.RowCount; vLoop++)
            {
                //ISG_BOX_LIST.SetCellValue(vLoop, 0, ICB_ALL.CheckBoxValue.ToString());
                //IDA_DISPOSAL_LINE.OraSelectData.AcceptChanges();
                //IDA_DISPOSAL_LINE.Refillable = true;

                if (ISG_BOX_LIST.GetCellValue(vLoop, 0).ToString() == "N")
                {
                    ISG_BOX_LIST.SetCellValue(vLoop, 0, ICB_ALL.CheckBoxValue.ToString());
                    if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
                    {
                        IDC_FIFO_CHECK.SetCommandParamValue("W_INVENTORY_ITEM_ID", ISG_BOX_LIST.GetCellValue(vLoop, ISG_BOX_LIST.GetColumnToIndex("INVENTORY_ITEM_ID")));
                        IDC_FIFO_CHECK.SetCommandParamValue("W_BOX_NO", ISG_BOX_LIST.GetCellValue(vLoop, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO")));

                        IDC_FIFO_CHECK.ExecuteNonQuery();

                        string vErrMsg = Convert.ToString(IDC_FIFO_CHECK.GetCommandParamValue("X_ERR_MSG"));

                        if (vErrMsg != "OK")
                        {
                            MessageBoxAdv.Show(vErrMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ISG_BOX_LIST.SetCellValue(vLoop, ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG"), "N");
                            return;
                        }
                        ISGridAdvExChangedEventArgs vISGridAdvExChangedEventArgs = new ISGridAdvExChangedEventArgs(vLoop, 0, "N", "Y");
                        ISG_BOX_LIST_CurrentCellChanged(this, vISGridAdvExChangedEventArgs);
                    }
                }

                if (ISG_BOX_LIST.GetCellValue(vLoop, 0).ToString() == "Y")
                {
                    ISG_BOX_LIST.SetCellValue(vLoop, 0, ICB_ALL.CheckBoxValue.ToString());
                    if (e.CheckedState == ISUtil.Enum.CheckedState.Unchecked) 
                    {
                        ISGridAdvExChangedEventArgs vISGridAdvExChangedEventArgs = new ISGridAdvExChangedEventArgs(vLoop, 0, "Y", "N");
                        ISG_BOX_LIST_CurrentCellChanged(this, vISGridAdvExChangedEventArgs);
                    }
                }
            }


            //for (int i = 0; i < ISG_BOX_LIST.RowCount; i++)
            //{
            //    string vSelectFlag = Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG")));
            //    string vPackingBoxNo = Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO")));

            //    ISG_BOX_LIST.SetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG"), ICB_ALL.CheckBoxValue);

            //    if (vPackingBoxNo != "NOT_PACKING")
            //    {
            //        if (Convert.ToString(ICB_ALL.CheckBoxValue) == "Y")
            //        {
            //            if (vSelectFlag == "N")
            //            {
            //                for (int j = 0; j < ISG_ORDER_LIST.RowCount; j++)
            //                {
            //                    string vItemCode = Convert.ToString(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("ITEM_CODE")));
            //                    if (vItemCode == Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("ITEM_CODE"))))
            //                    {
            //                        decimal vDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) + Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("ONHAND_QTY")));
            //                        decimal vRemainDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("REAMAIN_DELIVERY_QTY")));

            //                        if (vRemainDeliveryQty < vDeliveryQty)
            //                        {
            //                            MessageBoxAdv.Show("지시잔량보다 출고수량이 많습니다.");
            //                            return;
            //                        }

            //                        ISG_ORDER_LIST.SetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), vDeliveryQty);

            //                    }
            //                }
            //            }
            //        }
            //        if (Convert.ToString(ICB_ALL.CheckBoxValue) == "N")
            //        {
            //            if (vSelectFlag == "Y")
            //            {
            //                for (int j = 0; j < ISG_ORDER_LIST.RowCount; j++)
            //                {
            //                    string vItemCode = Convert.ToString(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("ITEM_CODE")));

            //                    if (vItemCode == Convert.ToString(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("ITEM_CODE"))))
            //                    {
            //                        if (vPackingBoxNo == "NOT_PACKING")
            //                        {
            //                            ISG_ORDER_LIST.SetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) - Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("PACKING_QTY"))));
            //                        }
            //                        else
            //                        {
            //                            ISG_ORDER_LIST.SetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(j, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) - Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("ONHAND_QTY"))));
            //                        }
            //                    }
            //                }
            //            }
            //        }

            //        //ISG_BOX_LIST.SetCellValue(i, ISG_BOX_LIST.GetColumnToIndex("SELECT_FLAG"), ICB_ALL.CheckBoxValue);
            //    }

            //    if (e.CheckedState == ISUtil.Enum.CheckedState.Checked)
            //    {
            //        ISGridAdvExChangedEventArgs vISGridAdvExChangedEventArgs = new ISGridAdvExChangedEventArgs(i, 0, "N", "Y");
            //        ISG_BOX_LIST_CurrentCellChanged(this, vISGridAdvExChangedEventArgs);
            //    }
            //    else
            //    {
            //        ISGridAdvExChangedEventArgs vISGridAdvExChangedEventArgs = new ISGridAdvExChangedEventArgs(i, 0, "Y", "N");
            //        ISG_BOX_LIST_CurrentCellChanged(this, vISGridAdvExChangedEventArgs);
            //    }
            //}
        }

        private void ISG_BOX_LIST_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            decimal vDeliveryQty = 0;

            switch (ISG_BOX_LIST.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "SELECT_FLAG":
                    if (Convert.ToString(ISG_BOX_LIST.GetCellValue(e.RowIndex, e.ColIndex)) == "Y")
                    {
                        for (int i = 0; i < ISG_ORDER_LIST.RowCount; i++)
                        {
                            string vItemCode = Convert.ToString(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("ITEM_CODE")));
                            if (vItemCode == Convert.ToString(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("ITEM_CODE"))))
                            {
                                IDC_FIFO_CHECK.SetCommandParamValue("W_INVENTORY_ITEM_ID", ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("INVENTORY_ITEM_ID")));
                                IDC_FIFO_CHECK.SetCommandParamValue("W_BOX_NO", ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO")));

                                IDC_FIFO_CHECK.ExecuteNonQuery();

                                string vErrMsg = Convert.ToString(IDC_FIFO_CHECK.GetCommandParamValue("X_ERR_MSG"));

                                if (vErrMsg != "OK")
                                {
                                    MessageBoxAdv.Show(vErrMsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    ISG_BOX_LIST.SetCellValue(e.RowIndex, e.ColIndex, "N");
                                    return;
                                }

                                decimal vOrderRemainDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("REAMAIN_DELIVERY_QTY")));
                                decimal vOrderDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY")));
                                decimal vPackingQty = Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("ONHAND_QTY")));

                                if (Convert.ToString(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO"))) == "NOT_PACKING")
                                {
                                    if (vOrderRemainDeliveryQty <= vPackingQty)
                                    {
                                        ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), vOrderRemainDeliveryQty - vOrderDeliveryQty);
                                    }
                                    else
                                    {
                                        ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), vPackingQty);
                                    }
                                }
                                else
                                {
                                    ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("ONHAND_QTY")));
                                }

                                vDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) + Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY")));

                                if (vOrderRemainDeliveryQty < vDeliveryQty)
                                {
                                    if (Convert.ToString(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO"))) == "NOT_PACKING")
                                    {
                                        ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), vOrderRemainDeliveryQty - vOrderDeliveryQty);
                                        ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), vOrderRemainDeliveryQty);
                                    }
                                    else
                                    {
                                        //출고수량이 지시잔량 보다 많습니다.
                                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10065").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                        ISG_BOX_LIST.SetCellValue(e.RowIndex, e.ColIndex, "N");
                                        ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), 0);
                                        return;
                                    }
                                }
                                else
                                {
                                    ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), Convert.ToDecimal(vDeliveryQty));                                    
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < ISG_ORDER_LIST.RowCount; i++)
                        {
                            string vItemCode = Convert.ToString(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("ITEM_CODE")));

                            if (vItemCode == Convert.ToString(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("ITEM_CODE"))))
                            {
                                //if (Convert.ToString(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO"))) == "NOT_PACKING")
                                //{
                                //    ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) - Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"))));
                                //}
                                //else
                                //{
                                //    ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) - Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("ONHAND_QTY"))));
                                //}
                                ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) - Convert.ToDecimal(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"))));

                                ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), 0);
                            }
                        }
                    }

                    break;

                default:
                    break;
            }
        }



        private void Set_GRID_STATUS_ROW(object pRESPONSE_DESC)
        {
            int vSTATUS = 0;                // INSERTABLE, UPDATABLE;

            int vROW = ISG_BOX_LIST.RowIndex;

            if (iString.ISNull(pRESPONSE_DESC) == "NOT_PACKING")
            {
                vSTATUS = 1;
            }

            ISG_BOX_LIST.GridAdvExColElement[ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY")].Insertable = vSTATUS;
            ISG_BOX_LIST.GridAdvExColElement[ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY")].Updatable = vSTATUS;

        }

        private void IDA_BOX_LIST_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {
            if (pBindingManager.DataRow == null)
            {
                return;
            }
            Set_GRID_STATUS_ROW(pBindingManager.DataRow["PACKING_BOX_NO"]);
        }

        private void ISG_BOX_LIST_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            decimal vDeliveryQty = 0;

            switch (ISG_BOX_LIST.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "DELIVERY_QTY":

                    for (int i = 0; i < ISG_ORDER_LIST.RowCount; i++)
                    {
                        string vItemCode = Convert.ToString(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("ITEM_CODE")));
                        if (vItemCode == Convert.ToString(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("ITEM_CODE"))))
                        {
                            ////////////////////
                            if (Convert.ToDecimal(e.NewValue) > 5)
                            {
                                object test11 = ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"));
                            }
                            ////////////////////
                            decimal vOrderRemainDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("REAMAIN_DELIVERY_QTY")));
                            decimal vOrderDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY")));

                            ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), vOrderDeliveryQty - iString.ISNumtoZero(e.OldValue));

                            //if (Convert.ToString(ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("PACKING_BOX_NO"))) == "NOT_PACKING")
                            //{
                            //    ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), vOrderRemainDeliveryQty - vOrderDeliveryQty);
                            //}
                            //else
                            //{
                            //    ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), ISG_BOX_LIST.GetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("ONHAND_QTY")));
                            //}

                            vDeliveryQty = Convert.ToDecimal(ISG_ORDER_LIST.GetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"))) + Convert.ToDecimal(e.NewValue);

                            if (vOrderRemainDeliveryQty < vDeliveryQty)
                            {
                                //출고수량이 지시잔량 보다 많습니다.
                                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("OE_10065").ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //ISG_BOX_LIST.SetCellValue(e.RowIndex, e.ColIndex, "N");
                                ISG_BOX_LIST.SetCellValue(e.RowIndex, ISG_BOX_LIST.GetColumnToIndex("DELIVERY_QTY"), e.OldValue);
                                return;
                            }
                            ISG_ORDER_LIST.SetCellValue(i, ISG_ORDER_LIST.GetColumnToIndex("DELIVERY_QTY"), Convert.ToDecimal(vDeliveryQty));
                        }
                    }

                    break;

                default:
                    break;
            }
        }        

    }

}
