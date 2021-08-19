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

namespace HRMF0392
{
    public partial class HRMF0392 : Office2007Form
    {        
        #region ----- Variables -----

        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();
        ISFunction.ISConvert iConv = new ISFunction.ISConvert();

        object mDB_IP = string.Empty;
        object mDB_PORT = string.Empty;
        object mDB_NAME = string.Empty;
        object mDB_USER = string.Empty;
        object mDB_PWD = string.Empty;

        string mDB_TYPE = "ALARM";
        string mDEVICE_TYPE = "BIO";

        object mDAILY_FLAG = "N"; 

        #endregion;

        #region ----- Constructor -----

        public HRMF0392()
        {
            InitializeComponent();
        }

        public HRMF0392(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- SQL DB 연결 / SQL DB DATA --> Oracle 저장 -----

        private Boolean Connect_SQL_DB(object pWORK_DATE_FR, object pWORK_DATE_TO)
        {
            SqlConnection mCONN = new SqlConnection();

            DataSet vDataSet = null;
            SqlDataAdapter vAdater = null;

            string vCON_SQL = string.Empty;
            string vQUERY_STRING = string.Empty;

            string vSTATUS = "F";
            string vMESSAGE = null;

            //MessageBoxAdv.Show(string.Format("2, {0}-{1}", pWORK_DATE_FR, pWORK_DATE_TO));

            ipbSECOM_INTERFACE.BarFillPercent = 0;
            try
            {
                vCON_SQL = string.Format(@"Server={0},{1};DataBase={2};UID={3};PWD={4}", mDB_IP, mDB_PORT, mDB_NAME, mDB_USER, mDB_PWD);
                mCONN.ConnectionString = vCON_SQL;
                mCONN.Open();

                vQUERY_STRING = "  SELECT EI.USERID ";
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.CHECK_DATETIME");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.CHECK_DATE ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.nDateTime"); 
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.nEventIdn");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.EVENT_TYPE ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.EVENT_TYPE_DESC ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.CARD_NUM ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.NAME ");
                //vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.nIsLog ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.EVENT_STATUS ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.EVENT_STATUS_DESC ");
                //vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.DEVICE_PORT ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.DEVICE_ID ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.READER_NAME ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.READER_CODE ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.DEVICE_IP ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.PERSON_NUM ");
                vQUERY_STRING = string.Format("{0} {1}", vQUERY_STRING, "FROM TB_EVENT_ERP_INTERFACE EI ");
                vQUERY_STRING = string.Format("{0} WHERE EI.CHECK_DATE BETWEEN '{1}' AND '{2}' ", vQUERY_STRING, pWORK_DATE_FR, pWORK_DATE_TO);
                vQUERY_STRING = string.Format("{0} ORDER BY EI.CHECK_DATETIME;", vQUERY_STRING);
                
                //CAST(CONVERT(nVARCHAR(10), DATEADD(S, EL.nDateTime, '1970-01-01'), 120) AS DATETIME

                SqlCommand cmdString = new SqlCommand();
                cmdString.Connection = mCONN;
                cmdString.CommandType = CommandType.Text;
                cmdString.CommandTimeout = 10;
                cmdString.CommandText = vQUERY_STRING;

                vAdater = new System.Data.SqlClient.SqlDataAdapter();
                vAdater.SelectCommand = cmdString;

                vDataSet = new DataSet();
                vAdater.Fill(vDataSet, "TB_EVENT_V");

                DataTable vDataTable = new DataTable();
                vDataTable = vDataSet.Tables["TB_EVENT_V"];

                // insert.
                int vRowCount = 0;
                DateTime vSysDate = DateTime.Now;
                foreach (DataRow vRow in vDataTable.Rows)
                {
                    vRowCount = vRowCount + 1;
                    ipbSECOM_INTERFACE.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(vDataTable.Rows.Count)) * 100F;
                    iptSET_MESSAGE.PromptText = string.Format("{0}-{1}", vRow["PERSON_NUM"], vRow["NAME"]);
                    this.Cursor = Cursors.WaitCursor;
                    Application.DoEvents();

                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_CHECK_DATETIME", vRow["CHECK_DATETIME"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_CARD_NUM", vRow["CARD_NUM"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_USER_NAME", vRow["NAME"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_CORP_CODE", CORP_ID_0.EditValue);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_PERSON_KEY", vRow["PERSON_NUM"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_TYPE", vRow["EVENT_TYPE"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_DESC", vRow["EVENT_TYPE_DESC"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_STATUS", vRow["EVENT_STATUS"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_STATUS_DESC", vRow["EVENT_STATUS_DESC"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_ID", vRow["DEVICE_ID"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_CODE", vRow["READER_CODE"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_NAME", vRow["READER_NAME"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_CREATION_DATE", vSysDate);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                    IDC_INSERT_DEVICE_LOG.ExecuteNonQuery();
                    vSTATUS = IDC_INSERT_DEVICE_LOG.GetCommandParamValue("O_STATUS").ToString();
                    vMESSAGE = iConv.ISNull(IDC_INSERT_DEVICE_LOG.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_INSERT_DEVICE_LOG.ExcuteError || vSTATUS == "F")
                    {
                        vDataSet.Dispose();
                        vAdater.Dispose();
                        cmdString.Dispose();

                        mCONN.Close();
                        mCONN.Dispose();
                        MessageBoxAdv.Show(vMESSAGE, "Errro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false; 
                    }
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

        private Boolean Connect_MSSQL(object pWORK_DATE_FR, object pWORK_DATE_TO)
        {
            SqlConnection mCONN = new SqlConnection();

            DataSet vDataSet = null;
            SqlDataAdapter vAdater = null;

            string vCON_SQL = string.Empty;
            string vQUERY_STRING = string.Empty;

            string vSTATUS = "F";
            string vMESSAGE = null;

            //MessageBoxAdv.Show(string.Format("2, {0}-{1}", pWORK_DATE_FR, pWORK_DATE_TO));

            ipbSECOM_INTERFACE.BarFillPercent = 0;
            try
            {
                vCON_SQL = string.Format(@"Server={0};DataBase={1};UID={2};PWD={3}", mDB_IP, mDB_NAME, mDB_USER, mDB_PWD);
                mCONN.ConnectionString = vCON_SQL;
                mCONN.Open();

                vQUERY_STRING = "  SELECT DISTINCT CIO.ParentId ";
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.CheckTime ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.CHECK_DATE ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.ManualInput ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.EmployeeCode ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.Description ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.MachineNumber ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.EmployeeCode AS CARD_NUM ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.IsGroup ");
                //vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.nIsLog ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.InOutMode ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.Workcode ");
                //vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.DEVICE_PORT ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.BranchCode ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.VerifyMode ");
                //vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " EI.DEVICE_IP ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CIO.EmployeeCode AS PERSON_NUM ");
                vQUERY_STRING = string.Format("{0} {1}", vQUERY_STRING, "FROM IF_HRM_CHECK_INOUT_V CIO ");
                vQUERY_STRING = string.Format("{0} WHERE CIO.CHECK_DATE BETWEEN '{1}' AND '{2}' ", vQUERY_STRING, pWORK_DATE_FR, pWORK_DATE_TO);
                vQUERY_STRING = string.Format("{0} ORDER BY CIO.CheckTime ;", vQUERY_STRING);

                SqlCommand cmdString = new SqlCommand();
                cmdString.Connection = mCONN;
                cmdString.CommandType = CommandType.Text;
                cmdString.CommandTimeout = 10;
                cmdString.CommandText = vQUERY_STRING;

                vAdater = new System.Data.SqlClient.SqlDataAdapter();
                vAdater.SelectCommand = cmdString;

                vDataSet = new DataSet();
                vAdater.Fill(vDataSet, "TB_EVENT_V");

                DataTable vDataTable = new DataTable();
                vDataTable = vDataSet.Tables["TB_EVENT_V"];

                // insert.
                int vRowCount = 0;
                DateTime vSysDate = DateTime.Now;
                foreach (DataRow vRow in vDataTable.Rows)
                {
                    vRowCount = vRowCount + 1;
                    ipbSECOM_INTERFACE.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(vDataTable.Rows.Count)) * 100F;
                    iptSET_MESSAGE.PromptText = string.Format("{0} / {1} *** {2} :: {3}", vRowCount, vDataTable.Rows.Count, vRow["CHECK_DATE"], vRow["PERSON_NUM"]);
                    Application.DoEvents();

                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_ZKTIME_ID", vRow["ParentId"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_EMPNO", vRow["PERSON_NUM"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_EMPNAME", vRow["PERSON_NUM"]);
                    //IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_EMPID", vRow["PERSON_NUM"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_MACHINECODE", vRow["MachineNumber"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_RECORDTYPE", vRow["InOutMode"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_BRUSHDATETIME", vRow["CheckTime"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_CHECK_DATE", vRow["CHECK_DATE"]);
                    //IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_SPPLYOPERATOR", vRow["SPPLYOPERATOR"]);
                    //IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_SPPLYREASON", vRow["SPPLYREASON"]);
                    //IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_SPPLYDATE", vRow["SPPLYDATE"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_CUSER1", vRow["VerifyMode"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_CUSER2", vRow["Workcode"]);
                    //IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_DEPTNO", vRow["DEPTNO"]);
                    //IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_DEPTNAME", vRow["DEPTNAME"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_PERSON_NUM", vRow["CARD_NUM"]);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                    IDC_DEVICE_HISTORY_INSERT.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                    IDC_DEVICE_HISTORY_INSERT.ExecuteNonQuery();
                    vSTATUS = iConv.ISNull(IDC_DEVICE_HISTORY_INSERT.GetCommandParamValue("O_STATUS"));
                    vMESSAGE = iConv.ISNull(IDC_DEVICE_HISTORY_INSERT.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_DEVICE_HISTORY_INSERT.ExcuteError || vSTATUS == "F")
                    {
                        vDataSet.Dispose();
                        vAdater.Dispose();
                        cmdString.Dispose();

                        mCONN.Close();
                        mCONN.Dispose();
                        MessageBoxAdv.Show(vMESSAGE, "Errro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
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

        private Boolean Connect_WISEEYE_DB(object pWORK_DATE_FR, object pWORK_DATE_TO)
        {
            SqlConnection mCONN = new SqlConnection();

            DataSet vDataSet = null;
            SqlDataAdapter vAdater = null;

            string vCON_SQL = string.Empty;
            string vQUERY_STRING = string.Empty;

            string vSTATUS = "F";
            string vMESSAGE = null;

            //MessageBoxAdv.Show(string.Format("2, {0}-{1}", pWORK_DATE_FR, pWORK_DATE_TO));

            ipbSECOM_INTERFACE.BarFillPercent = 0;
            try
            {
                vCON_SQL = string.Format(@"Server={0},{1};DataBase={2};UID={3};PWD={4}", mDB_IP, mDB_PORT, mDB_NAME, mDB_USER, mDB_PWD);
                mCONN.ConnectionString = vCON_SQL;
                mCONN.Open(); 

                vQUERY_STRING = "  SELECT CONVERT(VARCHAR(20), [NgayCham], 120) AS CHECK_DATETIME ";
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [MayCham] AS DEVICE_CODE ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " CONVERT(varchar(10), [NgayCham], 120) AS IO_DATE "); 
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Ngay] AS DAY ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Thang] AS MONTH ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Nam] AS YEAR ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Gio] AS HOUR ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [Phut] AS MIN ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " '55' AS EVENT_TYPE ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " '' AS EVENT_TYPE_DESC ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " '' AS EVENT_STATUS ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " '1' AS DEVICE_ID "); 
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [MaCham] AS CARD_NUM ");
                vQUERY_STRING = string.Format("{0}, {1}", vQUERY_STRING, " [MaCham] AS NAME ");
                vQUERY_STRING = string.Format("{0} {1}", vQUERY_STRING, " FROM [dbo].[TimeLog] ");
                vQUERY_STRING = string.Format("{0} WHERE CONVERT(varchar(10), [NgayCham], 120) BETWEEN '{1}' AND '{2}' ", vQUERY_STRING, pWORK_DATE_FR, pWORK_DATE_TO);
                vQUERY_STRING = string.Format("{0} ORDER BY [NgayCham];", vQUERY_STRING); 

                //CAST(CONVERT(nVARCHAR(10), DATEADD(S, EL.nDateTime, '1970-01-01'), 120) AS DATETIME

                SqlCommand cmdString = new SqlCommand();
                cmdString.Connection = mCONN;
                cmdString.CommandType = CommandType.Text;
                cmdString.CommandTimeout = 10;
                cmdString.CommandText = vQUERY_STRING;

                vAdater = new System.Data.SqlClient.SqlDataAdapter();
                vAdater.SelectCommand = cmdString;

                vDataSet = new DataSet();
                vAdater.Fill(vDataSet, "TB_EVENT_V");

                DataTable vDataTable = new DataTable();
                vDataTable = vDataSet.Tables["TB_EVENT_V"];

                // insert.
                int vRowCount = 0;
                DateTime vSysDate = DateTime.Now;
                foreach (DataRow vRow in vDataTable.Rows)
                {
                    vRowCount = vRowCount + 1;
                    ipbSECOM_INTERFACE.BarFillPercent = (Convert.ToSingle(vRowCount) / Convert.ToSingle(vDataTable.Rows.Count)) * 100F;
                    iptSET_MESSAGE.PromptText = string.Format("{0}-{1}", vRow["CARD_NUM"], vRow["NAME"]);
                    this.Cursor = Cursors.WaitCursor;
                    Application.DoEvents();

                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_CHECK_DATETIME", vRow["CHECK_DATETIME"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_CARD_NUM", vRow["CARD_NUM"]);
                    //idcINSERT_DEVICE_LOG.SetCommandParamValue("P_USER_NAME", vRow["NAME"]);
                    //idcINSERT_DEVICE_LOG.SetCommandParamValue("P_CORP_CODE", CORP_ID_0.EditValue);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_PERSON_KEY", vRow["CARD_NUM"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_TYPE", vRow["EVENT_TYPE"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_DESC", vRow["EVENT_TYPE_DESC"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_STATUS", vRow["EVENT_STATUS"]);
                    //idcINSERT_DEVICE_LOG.SetCommandParamValue("P_EVENT_STATUS_DESC", vRow["EVENT_STATUS_DESC"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_ID", vRow["DEVICE_ID"]);
                    //idcINSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_CODE", vRow["READER_CODE"]);
                    //idcINSERT_DEVICE_LOG.SetCommandParamValue("P_DEVICE_NAME", vRow["READER_NAME"]);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_CREATION_DATE", vSysDate);
                    IDC_INSERT_DEVICE_LOG.SetCommandParamValue("P_USER_ID", isAppInterfaceAdv1.USER_ID);
                    IDC_INSERT_DEVICE_LOG.ExecuteNonQuery();
                    vSTATUS = IDC_INSERT_DEVICE_LOG.GetCommandParamValue("O_STATUS").ToString();
                    vMESSAGE = iConv.ISNull(IDC_INSERT_DEVICE_LOG.GetCommandParamValue("O_MESSAGE"));
                    if (IDC_INSERT_DEVICE_LOG.ExcuteError || vSTATUS == "F")
                    {
                        vDataSet.Dispose();
                        vAdater.Dispose();
                        cmdString.Dispose();

                        mCONN.Close();
                        mCONN.Dispose();
                        MessageBoxAdv.Show(vMESSAGE, "Errro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
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

        private void GetDeviceConfig()
        {
            if (mDB_TYPE == "MDB")
            {
                IDC_GET_WORK_DEVICE_MDB_P.SetCommandParamValue("W_STD_DATE", END_DATE_0.EditValue);
                IDC_GET_WORK_DEVICE_MDB_P.ExecuteNonQuery();
                mDB_IP = IDC_GET_WORK_DEVICE_MDB_P.GetCommandParamValue("O_PATH");
                mDB_NAME = IDC_GET_WORK_DEVICE_MDB_P.GetCommandParamValue("O_MDB_NAME");
                mDB_USER = IDC_GET_WORK_DEVICE_MDB_P.GetCommandParamValue("O_USER_ID");
                mDB_PWD = IDC_GET_WORK_DEVICE_MDB_P.GetCommandParamValue("O_USER_PWD");
                mDAILY_FLAG = IDC_GET_WORK_DEVICE_MDB_P.GetCommandParamValue("O_DAILY_FLAG");
            }
            else
            {
                IDC_GET_WORK_DEVICE_SQL_P.SetCommandParamValue("W_STD_DATE", END_DATE_0.EditValue);
                IDC_GET_WORK_DEVICE_SQL_P.ExecuteNonQuery();
                mDB_IP = iConv.ISNull(IDC_GET_WORK_DEVICE_SQL_P.GetCommandParamValue("O_DB_IP"));
                mDB_PORT = iConv.ISNull(IDC_GET_WORK_DEVICE_SQL_P.GetCommandParamValue("O_DB_PORT"));
                mDB_NAME = iConv.ISNull(IDC_GET_WORK_DEVICE_SQL_P.GetCommandParamValue("O_DB_NAME"));
                mDB_USER = iConv.ISNull(IDC_GET_WORK_DEVICE_SQL_P.GetCommandParamValue("O_DB_USER"));
                mDB_PWD = iConv.ISNull(IDC_GET_WORK_DEVICE_SQL_P.GetCommandParamValue("O_DB_PWD"));
            }
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

            idaDEVICE_LOG.Fill();
            igrDEVICE_LOG.Focus();
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
        
        private void HRMF0392_Load(object sender, EventArgs e)
        {
            igbSET_INTERFACE.Visible = false; 
        }

        private void HRMF0392_Shown(object sender, EventArgs e)
        {
            START_DATE_0.EditValue = DateTime.Today;
            END_DATE_0.EditValue = DateTime.Today;

            DefaultCorporation();

            //출퇴근 장치 구분//
            IDC_GET_WORK_DEVICE_TYPE.SetCommandParamValue("W_STD_DATE", END_DATE_0.EditValue);
            IDC_GET_WORK_DEVICE_TYPE.ExecuteNonQuery();
            mDB_TYPE = iConv.ISNull(IDC_GET_WORK_DEVICE_TYPE.GetCommandParamValue("O_DB_TYPE"));
            mDEVICE_TYPE = iConv.ISNull(IDC_GET_WORK_DEVICE_TYPE.GetCommandParamValue("O_WORK_DEVICE_TYPE"));

            //DefaultSetFormReSize();             //[Child Form, Mdi Form에 맞게 ReSize]
            irbALL.CheckedState = ISUtil.Enum.CheckedState.Checked;
            igbSET_INTERFACE.Visible = false;
        }

        private void irb_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv isSTATUS = sender as ISRadioButtonAdv;
            STATUS_FLAG.EditValue = isSTATUS.RadioCheckedString;
        }

        private void btnSET_INTERFACE_ButtonClick(object pSender, EventArgs pEventArgs)
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

            //WORK DEVICE환경//
            GetDeviceConfig();

            string vSTATUS = "";
            string vMESSAGE = ""; 
            int RecordCount;

            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            //출퇴근 장치 구분//
            IDC_GET_WORK_DEVICE_TYPE.ExecuteNonQuery();
            mDB_TYPE = iConv.ISNull(IDC_GET_WORK_DEVICE_TYPE.GetCommandParamValue("O_DB_TYPE"));
            mDEVICE_TYPE = iConv.ISNull(IDC_GET_WORK_DEVICE_TYPE.GetCommandParamValue("O_WORK_DEVICE_TYPE"));

            IDC_COUNT_DEVICE_LOG.ExecuteNonQuery();
            RecordCount = Convert.ToInt32(IDC_COUNT_DEVICE_LOG.GetCommandParamValue("O_COUNT"));
            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            if (RecordCount > 0)
            {
                if (DialogResult.OK == MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10082"), "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))
                {
                    Application.UseWaitCursor = true;
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    Application.DoEvents();
                    
                    vSTATUS = "F";
                    vMESSAGE = null;

                    // 기존 자료 삭제.
                    IDC_DELETE_DEVICE_LOG.ExecuteNonQuery();
                    vSTATUS = IDC_DELETE_DEVICE_LOG.GetCommandParamValue("O_STATUS").ToString();
                    vMESSAGE = iConv.ISNull(IDC_DELETE_DEVICE_LOG.GetCommandParamValue("O_MESSAGE"));

                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    Application.DoEvents();

                    if (IDC_SET_INTERFACE.ExcuteError || vSTATUS == "F")
                    {
                        MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    return;
                }
            }
            
            // 동기화 처리.
            iptSET_MESSAGE.PromptTextElement[0].Default = null;
            igbSET_INTERFACE.Visible = true;
            Application.UseWaitCursor = true;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            Application.DoEvents();

            string vDate_Fr = string.Format("{0:yyyy-MM-dd}", iDate.ISGetDate(START_DATE_0.EditValue));
            string vDate_To = string.Format("{0:yyyy-MM-dd}", iDate.ISGetDate(END_DATE_0.EditValue));

            vSTATUS = "";
            vMESSAGE = "";
            if (mDEVICE_TYPE == "WISEEYE")
            {
                //MessageBoxAdv.Show(string.Format("1, {0}-{1}", vDate_Fr, vDate_To));
                //SQL DB 접속
                if (Connect_WISEEYE_DB(vDate_Fr, vDate_To) == false)
                {
                    igbSET_INTERFACE.Visible = false;

                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    Application.DoEvents();
                    return;
                }
            } 
            else
            {  
                //MessageBoxAdv.Show(string.Format("1, {0}-{1}", vDate_Fr, vDate_To));
                //SQL DB 접속
                if (Connect_SQL_DB(vDate_Fr, vDate_To) == false)
                {
                    igbSET_INTERFACE.Visible = false;

                    Application.UseWaitCursor = false;
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    Application.DoEvents();
                    return;
                }
            }

            // 실제 HRD_ATTEND_INTERFACE 에 INSERT 함. 
            IDC_SET_INTERFACE.ExecuteNonQuery();
            vSTATUS = IDC_SET_INTERFACE.GetCommandParamValue("O_STATUS").ToString();
            vMESSAGE = IDC_SET_INTERFACE.GetCommandParamValue("O_MESSAGE").ToString();

            igbSET_INTERFACE.Visible = false;

            Application.UseWaitCursor = false;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            Application.DoEvents();

            if (IDC_SET_INTERFACE.ExcuteError || vSTATUS == "F")
            { 
                MessageBoxAdv.Show(vMESSAGE, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } 
            MessageBoxAdv.Show(vMESSAGE, "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void PERSON_NUM_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SearchDB();
            }
        }

        private void NAME_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SearchDB();
            }
        }

        #endregion

    }
}