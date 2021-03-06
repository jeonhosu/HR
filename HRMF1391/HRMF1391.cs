using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Odbc;
using ISCommonUtil;

namespace HRMF1391
{
    public partial class HRMF1391 : Office2007Form
    {
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
         
        object mDB_IP = string.Empty;
        object mDB_PORT = string.Empty;
        object mDB_NAME = string.Empty;
        object mDB_USER = string.Empty;
        object mDB_PWD = string.Empty;
        object mDAILY_FLAG = "N";

        string mDB_TYPE = "MDB";
        string mDEVICE_TYPE = "SECOM";

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public HRMF1391()
        {
            InitializeComponent();
        }

        public HRMF1391(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- MDB 연결 / MDB 값 --> Oracle 저장 -----


        //private Boolean Connect_Secom_MSSQL(object pWORK_DATE_FR, object pWORK_DATE_TO)
        //{
        //    SqlConnection mCONN = new SqlConnection();

        //    DataSet vDataSet = null;
        //    SqlDataAdapter vAdater = null;

        //    string vCON_SQL = string.Empty;
        //    string vQUERY_STRING = string.Empty;

        //    string vSTATUS = "F";
        //    string vMESSAGE = null;

        //    //MessageBoxAdv.Show(string.Format("2, {0}-{1}", pWORK_DATE_FR, pWORK_DATE_TO));

        //    ipbSECOM_INTERFACE.BarFillPercent = 0;
        //    try
        //    {
        //        vCON_SQL = string.Format(@"Server={0};DataBase={1};UID={2};PWD={3}", mSVR_IP, mDB_NAME, mUSER_ID, mUSER_PWD);
        //        mCONN.ConnectionString = vCON_SQL;
        //        mCONN.Open();

        //        vQUERY_STRING = "  SELECT EI.USERID ";
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.CHECK_DATETIME");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.CHECK_DATE ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.nDateTime");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.nEventIdn");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.EVENT_TYPE ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.EVENT_TYPE_DESC ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.CARD_NUM ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.NAME ");
        //        //vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.nIsLog ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.EVENT_STATUS ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.EVENT_STATUS_DESC ");
        //        //vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.DEVICE_PORT ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.DEVICE_ID ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.READER_NAME ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.READER_CODE ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.DEVICE_IP ");
        //        vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.PERSON_NUM ");
        //        vQUERY_STRING = string.Format("{0} {1}", vQUERY_STRING, "FROM TB_EVENT_ERP_INTERFACE EI ");
        //        vQUERY_STRING = string.Format("{0} WHERE EI.CHECK_DATE BETWEEN '{1}' AND '{2}' ", vQUERY_STRING, pWORK_DATE_FR, pWORK_DATE_TO);
        //        vQUERY_STRING = string.Format("{0} ORDER BY EI.CHECK_DATETIME;", vQUERY_STRING);

        //        //CAST(CONVERT(nVARCHAR(10), DATEADD(S, EL.nDateTime, '1970-01-01'), 120) AS DATETIME

        //        SqlCommand cmdString = new SqlCommand();
        //        cmdString.Connection = mCONN;
        //        cmdString.CommandType = CommandType.Text;
        //        cmdString.CommandTimeout = 10;
        //        cmdString.CommandText = vQUERY_STRING;

        //        vAdater = new System.Data.SqlClient.SqlDataAdapter();
        //        vAdater.SelectCommand = cmdString;

        //        vDataSet = new DataSet();
        //        vAdater.Fill(vDataSet, "TB_EVENT_V");

        //        DataTable vDataTable = new DataTable();
        //        vDataTable = vDataSet.Tables["TB_EVENT_V"];

        //        // insert.
        //        int vRowCount = 0;
        //        DateTime vSysDate = DateTime.Now;
        //        foreach (DataRow vRow in vDataTable.Rows)
        //        {
        //            vRowCount = vRowCount + 1;
        //            ipbSECOM_INTERFACE.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(vDataTable.Rows.Count)) * 100F;
        //            iptSET_MESSAGE.PromptText = string.Format("{0}-{1}", vRow["PERSON_NUM"], vRow["NAME"]);
        //            Application.DoEvents();


        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_CHECK_DATETIME", vRow["CHECK_DATETIME"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_CARD_NUM", vRow["CARD_NUM"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_USER_NAME", vRow["NAME"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_CORP_CODE", CORP_ID_0.EditValue);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_PERSON_KEY", vRow["PERSON_NUM"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_TYPE", vRow["EVENT_TYPE"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_DESC", vRow["EVENT_TYPE_DESC"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_STATUS", vRow["EVENT_STATUS"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_STATUS_DESC", vRow["EVENT_STATUS_DESC"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_ID", vRow["DEVICE_ID"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_CODE", vRow["READER_CODE"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_NAME", vRow["READER_NAME"]);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_CREATION_DATE", vSysDate);
        //            idcINSERT_DEVICE_LOG.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
        //            idcINSERT_DEVICE_LOG.ExecuteNonQuery();
        //            vSTATUS = idcINSERT_DEVICE_LOG.GetCommandParamValue("O_STATUS").ToString();
        //            vMESSAGE = iString.ISNull(idcINSERT_DEVICE_LOG.GetCommandParamValue("O_MESSAGE"));
        //            if (idcINSERT_DEVICE_LOG.ExcuteError || vSTATUS == "F")
        //            {
        //                vDataSet.Dispose();
        //                vAdater.Dispose();
        //                cmdString.Dispose();

        //                mCONN.Close();
        //                mCONN.Dispose();
        //                MessageBoxAdv.Show(vMESSAGE, "Errro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                return false;
        //            }
        //        }

        //        vDataSet.Dispose();
        //        vAdater.Dispose();
        //        cmdString.Dispose();

        //        mCONN.Close();
        //        mCONN.Dispose();
        //        return true;
        //    }
        //    catch (System.Exception ex)
        //    {
        //        if (mCONN.State == ConnectionState.Open)
        //        {
        //            mCONN.Close();
        //            mCONN.Dispose();
        //        }
        //        MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return false;
        //    }
        //}


        private bool Connect_Secom_MSSQL(object pWORK_DATE_FR, object pWORK_DATE_TO)
        {
            //MSSQL 처리//
            ipbSECOM_INTERFACE.BarFillPercent = 0;

            SqlConnection mCONN = new SqlConnection();

            DataSet vDataSet = null;
            SqlDataAdapter vAdater = null;

            string vCON_SQL = string.Empty;
            string vQUERY_STRING = string.Empty;

            string vSTATUS = "F";
            string vMESSAGE = null;

            //MessageBoxAdv.Show(string.Format("2, {0}-{1}", pWORK_DATE_FR, pWORK_DATE_TO)); 
            try
            {
                vCON_SQL = string.Format(@"Server={0},{1};DataBase={2};UID={3};PWD={4}", mDB_IP, mDB_PORT, mDB_NAME, mDB_USER, mDB_PWD);
                mCONN.ConnectionString = vCON_SQL;
                mCONN.Open();

                vQUERY_STRING = "  SELECT [ATime] AS ATIME ";
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [ID] AS ID_SEQ ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [EqCode] AS EQCODE_A ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Master] AS MASTER_A ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Param] AS PARAM_A ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Ack] AS ACK ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [AckUser] AS ACKUSER ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [AckTime] AS ACKTIME ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [AckContent] AS ACKCONTENT "); 
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Transfer] AS TRANSFER ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [AckMode] AS MODE_A "); 
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [CardNo] AS CARDNO ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Name] AS CONTENT_A ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [State] + '-' + [Flag1]  + [Flag2] + [Flag3] + [Flag4] AS STATE "); 
                vQUERY_STRING = string.Format("{0} {1}", vQUERY_STRING, " FROM [dbo].[T_SECOM_ALARM] ");
                vQUERY_STRING = string.Format("{0} WHERE LEFT([ATime], 8) BETWEEN '{1}' AND '{2}' ", vQUERY_STRING, pWORK_DATE_FR, pWORK_DATE_TO);
                vQUERY_STRING = string.Format("{0} ORDER BY ATime;", vQUERY_STRING);

                //CAST(CONVERT(nVARCHAR(10), DATEADD(S, EL.nDateTime, '1970-01-01'), 120) AS DATETIME

                SqlCommand cmdString = new SqlCommand();
                cmdString.Connection = mCONN;
                cmdString.CommandType = CommandType.Text;
                cmdString.CommandTimeout = 10;
                cmdString.CommandText = vQUERY_STRING;

                vAdater = new System.Data.SqlClient.SqlDataAdapter();
                vAdater.SelectCommand = cmdString;

                vDataSet = new DataSet();
                vAdater.Fill(vDataSet, "ALARM");

                DataTable vDataTable = new DataTable();
                vDataTable = vDataSet.Tables["ALARM"];

                // insert.
                int vRowCount = 0;
                DateTime vSysDate = DateTime.Now;
                foreach (DataRow vRow in vDataTable.Rows)
                {
                    vRowCount = vRowCount + 1;
                    
                    ipbSECOM_INTERFACE.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(vDataTable.Rows.Count)) * 100F;
                    iptPERIOD.PromptText = string.Format("Set Interface Date : {0}~{1}", pWORK_DATE_FR, pWORK_DATE_TO);
                    iptSET_MESSAGE.PromptText = string.Format("{0}-{1}", vRow["CARDNO"], vRow["CONTENT_A"]);

                    Application.DoEvents();  

                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ATIME", vRow["ATIME"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ID_SEQ", vRow["ID_SEQ"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_EQCODE_A", vRow["EQCODE_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_MASTER_A", vRow["MASTER_A"]);
                    //IDC_SECOM_HISTORY_INSERT.SetCommandParamValue("P_LOCAL_A", vRow["LOCAL_A"]);
                    //IDC_SECOM_HISTORY_INSERT.SetCommandParamValue("P_POINT_A", vRow["POINT_A"]);
                    //IDC_SECOM_HISTORY_INSERT.SetCommandParamValue("P_LOOP_A", vRow["LOOP_A"]);
                    //IDC_SECOM_HISTORY_INSERT.SetCommandParamValue("P_EQNAME", vRow["EQNAME"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_STATE", vRow["STATE"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_PARAM_A", vRow["PARAM_A"]);
                    //IDC_SECOM_HISTORY_INSERT.SetCommandParamValue("P_USER_A", vRow["USER_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_CONTENT_A", vRow["CONTENT_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ACK", vRow["ACK"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ACKUSER", vRow["ACKUSER"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ACKCONTENT", vRow["ACKCONTENT"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ACKTIME", vRow["ACKTIME"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_TRANSFER", vRow["TRANSFER"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_MODE_A", vRow["MODE_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_CARDNO", vRow["CARDNO"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_CREATION_DATE", vSysDate);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                    IDC_SET_VINA_TO_ERP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_VINA_TO_ERP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_VINA_TO_ERP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_VINA_TO_ERP.ExcuteError || vSTATUS == "F")
                    {
                        vDataSet.Dispose();
                        vAdater.Dispose();
                        cmdString.Dispose();

                        mCONN.Close();
                        mCONN.Dispose();
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false; 
                    } 
                    ipbSECOM_INTERFACE.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(vDataTable.Rows.Count)) * 100F; 
                }

                vDataSet.Dispose();
                vAdater.Dispose();
                cmdString.Dispose();

                mCONN.Close();
                mCONN.Dispose();
                return true;
            }
            catch (System.Exception ex)
            {
                if (mCONN.State == ConnectionState.Open)
                {
                    mCONN.Close();
                    mCONN.Dispose();
                }
                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }          
        }

        private bool Connect_Secom_MDB(string pWork_Date)
        {
            ipbSECOM_INTERFACE.BarFillPercent = 0;

            System.Data.DataTable mDataTable = null;
            System.Data.OleDb.OleDbDataAdapter vOleDataAdapter = null;
            System.Data.OleDb.OleDbCommand vOleCommand = null;
            System.Data.OleDb.OleDbConnection vOleConnection = new System.Data.OleDb.OleDbConnection();

            try
            {
                //secom data 폴더에 읽기/쓰기 권한이 있어야 함
                //당일 데이터 동기화 할때 오류 방지
                string vMDB = iConv.ISNull(mDB_NAME).Replace("YYYYMMDD", pWork_Date);
                vMDB = string.Format("{0}{1}", mDB_IP, vMDB);  
                string vConnectString = string.Format("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = {0};", vMDB);

                vOleConnection.ConnectionString = vConnectString; 
                vOleConnection.Open();
            }
            catch(System.Exception ex)
            {
                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            try
            { 
                System.Text.StringBuilder vQueryString = new System.Text.StringBuilder();

                vQueryString.Append("  SELECT [ALARM.ATIME] AS ATIME ");
                vQueryString.Append("       , [ALARM.ID] AS ID_SEQ ");
                vQueryString.Append("       , [ALARM.EQCODE] AS EQCODE_A ");
                vQueryString.Append("       , [ALARM.MASTER] AS MASTER_A ");
                vQueryString.Append("       , [ALARM.LOCAL] AS LOCAL_A ");
                vQueryString.Append("       , [ALARM.POINT] AS POINT_A ");
                vQueryString.Append("       , [ALARM.LOOP] AS LOOP_A ");
                vQueryString.Append("       , [ALARM.EQNAME] AS EQNAME ");
                vQueryString.Append("       , [ALARM.STATE] AS STATE ");
                vQueryString.Append("       , [ALARM.PARAM] AS PARAM_A ");
                vQueryString.Append("       , [ALARM.USER] AS USER_A ");
                vQueryString.Append("       , [ALARM.CONTENT] AS CONTENT_A ");
                vQueryString.Append("       , [ALARM.ACK] AS ACK ");
                vQueryString.Append("       , [ALARM.ACKUSER] AS ACKUSER ");
                vQueryString.Append("       , [ALARM.ACKCONTENT] AS ACKCONTENT ");
                vQueryString.Append("       , [ALARM.ACKTIME] AS ACKTIME ");
                vQueryString.Append("       , [ALARM.TRANSFER] AS TRANSFER");
                vQueryString.Append("       , [ALARM.MODE] AS MODE_A ");
                vQueryString.Append("       , [ALARM.CARDNO] AS CARDNO ");
                vQueryString.Append("    FROM ALARM ");
                vQueryString.Append("   WHERE LEFT(ALARM.ATIME, 8) =  '").Append(pWork_Date).Append("' ");
                vQueryString.Append("  ORDER BY ALARM.ATIME; ");
                
                vOleCommand = new System.Data.OleDb.OleDbCommand();
                vOleCommand.CommandType = System.Data.CommandType.Text;
                vOleCommand.CommandText = vQueryString.ToString();

                vOleCommand.Connection = vOleConnection;

                vOleDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
                vOleDataAdapter.SelectCommand = vOleCommand;

                mDataTable = new System.Data.DataTable();

                vOleDataAdapter.Fill(mDataTable);

                // insert.
                int vRowCount = 0;
                string vSTATUS = "";
                string vMESSAGE = "";
                DateTime vSysDate = DateTime.Now;
                foreach (System.Data.DataRow vRow in mDataTable.Rows)
                {
                    vRowCount = vRowCount + 1;
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ATIME", vRow["ATIME"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ID_SEQ", vRow["ID_SEQ"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_EQCODE_A", vRow["EQCODE_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_MASTER_A", vRow["MASTER_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_LOCAL_A", vRow["LOCAL_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_POINT_A", vRow["POINT_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_LOOP_A", vRow["LOOP_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_EQNAME", vRow["EQNAME"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_STATE", vRow["STATE"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_PARAM_A", vRow["PARAM_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_USER_A", vRow["USER_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_CONTENT_A", vRow["CONTENT_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ACK", vRow["ACK"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ACKUSER", vRow["ACKUSER"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ACKCONTENT", vRow["ACKCONTENT"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ACKTIME", vRow["ACKTIME"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_TRANSFER", vRow["TRANSFER"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_MODE_A", vRow["MODE_A"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_CARDNO", vRow["CARDNO"]);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_CREATION_DATE", vSysDate);
                    IDC_SET_VINA_TO_ERP.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                    IDC_SET_VINA_TO_ERP.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_SET_VINA_TO_ERP.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_SET_VINA_TO_ERP.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_SET_VINA_TO_ERP.ExcuteError)
                    {
                        MessageBoxAdv.Show(IDC_SET_VINA_TO_ERP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    else if (vSTATUS == "F")
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    ipbSECOM_INTERFACE.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(mDataTable.Rows.Count)) * 100F;
                }

                vOleDataAdapter.Dispose();
                vOleCommand.Dispose();
                mDataTable.Dispose();
                vOleConnection.Close();
                vOleConnection.Dispose();
            }
            catch (System.Exception ex)
            {
                vOleConnection.Close();
                vOleConnection.Dispose();
                
                MessageBoxAdv.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return true;
        }
          
        #endregion

        #region ----- Private Methods -----

        private void DefaultCorporation()
        {
            // Lookup SETTING
            ildCORP.SetLookupParamValue("W_DUTY_CONTROL_YN", "Y");
            ildCORP.SetLookupParamValue("W_ENABLED_FLAG_YN", "Y");

            // LOOKUP DEFAULT VALUE SETTING - CORP
            idcDEFAULT_CORP.SetCommandParamValue("W_DUTY_CONTROL_YN", "Y");
            idcDEFAULT_CORP.SetCommandParamValue("W_ENABLED_FLAG_YN", "Y");
            idcDEFAULT_CORP.ExecuteNonQuery();
            CORP_NAME_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_NAME");
            CORP_ID_0.EditValue = idcDEFAULT_CORP.GetCommandParamValue("O_CORP_ID");

            CORP_NAME_0.BringToFront();
        }
         

        private void SearchDB()
        {
            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iConv.ISNull(START_DATE_0.EditValue) == string.Empty)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (iConv.ISNull(END_DATE_0.EditValue) == string.Empty)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }

            if (Convert.ToDateTime(START_DATE_0.EditValue) > Convert.ToDateTime(END_DATE_0.EditValue))
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            IDA_VINA_HISTORY.Fill();
            IGR_VINA_HISTORY.Focus(); 
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchDB();
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

        #region ----- Form Event ----- 
        
        private void HRMF1391_Load(object sender, EventArgs e)
        {
            START_DATE_0.EditValue = DateTime.Today;
            END_DATE_0.EditValue = DateTime.Today;

            DefaultCorporation();
             
            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
            irbALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            igbSET_INTERFACE.Visible = false;            
        }

        private void irb_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv isSTATUS = sender as ISRadioButtonAdv;
            STATUS_FLAG.EditValue = isSTATUS.RadioCheckedString;
        }

        private void NAME_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SearchDB();
            }
        }

        private void START_DATE_0_CurrentEditValidating(object pSender, ISEditAdvValidatingEventArgs e)
        {
            END_DATE_0.EditValue = START_DATE_0.EditValue;
        }

        private void ibtnSET_SECOM_HISTORY_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            int RecordCount;

            if (iConv.ISNull(CORP_ID_0.EditValue) == string.Empty)
            {// 업체.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10001"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CORP_NAME_0.Focus();
                return;
            }
            if (iConv.ISNull(START_DATE_0.EditValue) == string.Empty)
            {// 시작일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10010"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                START_DATE_0.Focus();
                return;
            }
            if (iConv.ISNull(END_DATE_0.EditValue) == string.Empty)
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10011"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }

            if (Convert.ToDateTime(START_DATE_0.EditValue) > Convert.ToDateTime(END_DATE_0.EditValue))
            {// 종료일자
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10012"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                END_DATE_0.Focus();
                return;
            }
             
            string vSTATUS = "";
            string vMESSAGE = "";

            IDC_VINA_HISTORY_COUNT.ExecuteNonQuery();
            RecordCount = iConv.ISNumtoZero(IDC_VINA_HISTORY_COUNT.GetCommandParamValue("O_COUNT"));
            if (RecordCount > 0)
            {
                if (DialogResult.OK == MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))
                {
                    // 기존 자료 삭제.
                    IDC_VINA_HISTORY_DELETE.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_VINA_HISTORY_DELETE.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_VINA_HISTORY_DELETE.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_VINA_HISTORY_DELETE.ExcuteError)
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(IDC_VINA_HISTORY_DELETE.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else if (vSTATUS == "F")
                    {
                        Application.UseWaitCursor = false;
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        Application.DoEvents();

                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            Application.UseWaitCursor = true;
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            //
            IDC_SET_VINA_TO_ERP.ExecuteNonQuery();
            vSTATUS = iConv.ISNull(IDC_SET_VINA_TO_ERP.GetCommandParamValue("O_STATUS"));
            vMESSAGE = iConv.ISNull(IDC_SET_VINA_TO_ERP.GetCommandParamValue("O_MESSAGE"));
            if (IDC_SET_VINA_TO_ERP.ExcuteError)
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(IDC_SET_VINA_TO_ERP.ExcuteErrorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (vSTATUS == "F")
            {
                Application.UseWaitCursor = false;
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                Application.DoEvents();

                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Application.UseWaitCursor = false;
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            Application.DoEvents();

            MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        #endregion

    }
}