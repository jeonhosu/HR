namespace HRMF0631
{
    partial class HRMF0631_UPLOAD
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement11 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement1 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo oraConnectionInfo1 = new InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement10 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement6 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement7 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement8 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement9 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement1 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement2 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement3 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement4 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement5 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement6 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement7 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement8 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement9 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement10 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement11 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement12 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            this.V_STD_DATE = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.V_SYSDATE = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.V_CORP_ID = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.GB_UPLOAD_FILE = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.V_MESSAGE = new InfoSummit.Win.ControlAdv.ISPrompt();
            this.V_PB_INTERFACE = new InfoSummit.Win.ControlAdv.ISProgressBar();
            this.V_START_ROW = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.UPLOAD_FILE_PATH = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.BTN_CLOSED = new InfoSummit.Win.ControlAdv.ISButton();
            this.BTN_SELECT_EXCEL_FILE = new InfoSummit.Win.ControlAdv.ISButton();
            this.BTN_FILE_UPLOAD = new InfoSummit.Win.ControlAdv.ISButton();
            this.isPrompt1 = new InfoSummit.Win.ControlAdv.ISPrompt();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.IDC_IMPORT_EXCEL = new InfoSummit.Win.ControlAdv.ISDataCommand(this.components);
            this.IDC_GET_LOCAL_DATETIME_P = new InfoSummit.Win.ControlAdv.ISDataCommand(this.components);
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GB_UPLOAD_FILE)).BeginInit();
            this.GB_UPLOAD_FILE.SuspendLayout();
            // 
            // V_STD_DATE
            // 
            this.V_STD_DATE.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.V_STD_DATE.AutoScroll = true;
            this.V_STD_DATE.ComboBoxValue = "";
            this.V_STD_DATE.ComboData = null;
            this.V_STD_DATE.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_STD_DATE.DataAdapter = null;
            this.V_STD_DATE.DataColumn = null;
            this.V_STD_DATE.DateTimeValue = new System.DateTime(2020, 4, 22, 0, 0, 0, 0);
            this.V_STD_DATE.DoubleValue = 0D;
            this.V_STD_DATE.EditAdvType = InfoSummit.Win.ControlAdv.ISUtil.Enum.EditAdvType.DateTimeEdit;
            this.V_STD_DATE.EditValue = null;
            // 
            // HRMF0631_UPLOAD
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(716, 204);
            this.ControlBox = false;
            this.Controls.Add(this.GB_UPLOAD_FILE);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Name = "HRMF0631_UPLOAD";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Import File";
            this.Shown += new System.EventHandler(this.HRMF0631_UPLOAD_Shown);
            this.V_STD_DATE.Insertable = false;
            this.V_STD_DATE.Location = new System.Drawing.Point(8, 29);
            this.V_STD_DATE.LookupAdapter = null;
            this.V_STD_DATE.Name = "V_STD_DATE";
            this.V_STD_DATE.Nullable = true;
            this.V_STD_DATE.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_STD_DATE.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_STD_DATE.PromptText = "Std. Date";
            isLanguageElement11.Default = "Std. Date";
            isLanguageElement11.SiteName = null;
            isLanguageElement11.TL1_KR = "기준 일자";
            isLanguageElement11.TL2_CN = null;
            isLanguageElement11.TL3_VN = null;
            isLanguageElement11.TL4_JP = null;
            isLanguageElement11.TL5_XAA = null;
            this.V_STD_DATE.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement11});
            this.V_STD_DATE.PromptWidth = 120;
            this.V_STD_DATE.ReadOnly = true;
            this.V_STD_DATE.Size = new System.Drawing.Size(254, 21);
            this.V_STD_DATE.TabIndex = 197;
            this.V_STD_DATE.TabStop = false;
            this.V_STD_DATE.TextValue = "";
            this.V_STD_DATE.Updatable = false;
            // 
            // isAppInterfaceAdv1
            // 
            this.isAppInterfaceAdv1.AppMainButtonClick += new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv.ButtonEventHandler(this.isAppInterfaceAdv1_AppMainButtonClick);
            // 
            // V_SYSDATE
            // 
            this.V_SYSDATE.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.V_SYSDATE.AutoScroll = true;
            this.V_SYSDATE.ComboBoxValue = "";
            this.V_SYSDATE.ComboData = null;
            this.V_SYSDATE.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_SYSDATE.DataAdapter = null;
            this.V_SYSDATE.DataColumn = null;
            this.V_SYSDATE.DateFormat = "yyyy-MM-dd HH:mm:ss";
            this.V_SYSDATE.DateTimeValue = new System.DateTime(2020, 4, 22, 0, 0, 0, 0);
            this.V_SYSDATE.DoubleValue = 0D;
            this.V_SYSDATE.EditAdvType = InfoSummit.Win.ControlAdv.ISUtil.Enum.EditAdvType.DateTimeEdit;
            this.V_SYSDATE.EditValue = null;
            this.V_SYSDATE.Insertable = false;
            this.V_SYSDATE.Location = new System.Drawing.Point(268, 29);
            this.V_SYSDATE.LookupAdapter = null;
            this.V_SYSDATE.Name = "V_SYSDATE";
            this.V_SYSDATE.Nullable = true;
            this.V_SYSDATE.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_SYSDATE.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_SYSDATE.PromptText = "Sysdate";
            isLanguageElement1.Default = "Sysdate";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = "작업일시";
            isLanguageElement1.TL2_CN = null;
            isLanguageElement1.TL3_VN = null;
            isLanguageElement1.TL4_JP = null;
            isLanguageElement1.TL5_XAA = null;
            this.V_SYSDATE.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            this.V_SYSDATE.PromptWidth = 120;
            this.V_SYSDATE.ReadOnly = true;
            this.V_SYSDATE.Size = new System.Drawing.Size(280, 21);
            this.V_SYSDATE.TabIndex = 198;
            this.V_SYSDATE.TabStop = false;
            this.V_SYSDATE.TextValue = "";
            this.V_SYSDATE.Updatable = false;
            // 
            // V_CORP_ID
            // 
            this.V_CORP_ID.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.V_CORP_ID.AutoScroll = true;
            this.V_CORP_ID.ComboBoxValue = "";
            this.V_CORP_ID.ComboData = null;
            this.V_CORP_ID.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_CORP_ID.DataAdapter = null;
            this.V_CORP_ID.DataColumn = null;
            this.V_CORP_ID.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.V_CORP_ID.DoubleValue = 0D;
            this.V_CORP_ID.EditAdvType = InfoSummit.Win.ControlAdv.ISUtil.Enum.EditAdvType.NumberEdit;
            this.V_CORP_ID.EditValue = null;
            this.V_CORP_ID.Insertable = false;
            this.V_CORP_ID.Location = new System.Drawing.Point(571, 29);
            this.V_CORP_ID.LookupAdapter = null;
            this.V_CORP_ID.Name = "V_CORP_ID";
            this.V_CORP_ID.Nullable = true;
            this.V_CORP_ID.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_CORP_ID.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_CORP_ID.PromptText = "CORP_ID";
            isLanguageElement2.Default = "CORP_ID";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "CORP_ID";
            isLanguageElement2.TL2_CN = null;
            isLanguageElement2.TL3_VN = null;
            isLanguageElement2.TL4_JP = null;
            isLanguageElement2.TL5_XAA = null;
            this.V_CORP_ID.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.V_CORP_ID.PromptWidth = 90;
            this.V_CORP_ID.ReadOnly = true;
            this.V_CORP_ID.Size = new System.Drawing.Size(123, 21);
            this.V_CORP_ID.TabIndex = 196;
            this.V_CORP_ID.TabStop = false;
            this.V_CORP_ID.TextValue = "";
            this.V_CORP_ID.Updatable = false;
            this.V_CORP_ID.Visible = false;
            // 
            // isOraConnection1
            // 
            this.isOraConnection1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isOraConnection1.OraConnectionInfo = oraConnectionInfo1;
            this.isOraConnection1.OraHost = "1.241.249.174";
            this.isOraConnection1.OraPassword = "infoflex";
            this.isOraConnection1.OraPort = "1521";
            this.isOraConnection1.OraServiceName = "fekprod";
            this.isOraConnection1.OraUserId = "APPS";
            // 
            // isMessageAdapter1
            // 
            this.isMessageAdapter1.OraConnection = this.isOraConnection1;
            this.isMessageAdapter1.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            // 
            // GB_UPLOAD_FILE
            // 
            this.GB_UPLOAD_FILE.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.GB_UPLOAD_FILE.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.GB_UPLOAD_FILE.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.GB_UPLOAD_FILE.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.GB_UPLOAD_FILE.Controls.Add(this.V_SYSDATE);
            this.GB_UPLOAD_FILE.Controls.Add(this.V_STD_DATE);
            this.GB_UPLOAD_FILE.Controls.Add(this.V_CORP_ID);
            this.GB_UPLOAD_FILE.Controls.Add(this.V_MESSAGE);
            this.GB_UPLOAD_FILE.Controls.Add(this.V_PB_INTERFACE);
            this.GB_UPLOAD_FILE.Controls.Add(this.V_START_ROW);
            this.GB_UPLOAD_FILE.Controls.Add(this.UPLOAD_FILE_PATH);
            this.GB_UPLOAD_FILE.Controls.Add(this.BTN_CLOSED);
            this.GB_UPLOAD_FILE.Controls.Add(this.BTN_SELECT_EXCEL_FILE);
            this.GB_UPLOAD_FILE.Controls.Add(this.BTN_FILE_UPLOAD);
            this.GB_UPLOAD_FILE.Controls.Add(this.isPrompt1);
            this.GB_UPLOAD_FILE.Location = new System.Drawing.Point(4, 4);
            this.GB_UPLOAD_FILE.Name = "GB_UPLOAD_FILE";
            this.GB_UPLOAD_FILE.Padding = new System.Windows.Forms.Padding(5, 20, 5, 5);
            this.GB_UPLOAD_FILE.PromptText = "Select Upload file";
            isLanguageElement10.Default = "Select Upload file";
            isLanguageElement10.SiteName = null;
            isLanguageElement10.TL1_KR = "업로드 파일 선택";
            isLanguageElement10.TL2_CN = null;
            isLanguageElement10.TL3_VN = null;
            isLanguageElement10.TL4_JP = null;
            isLanguageElement10.TL5_XAA = null;
            this.GB_UPLOAD_FILE.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement10});
            this.GB_UPLOAD_FILE.Size = new System.Drawing.Size(708, 196);
            this.GB_UPLOAD_FILE.TabIndex = 7;
            // 
            // V_MESSAGE
            // 
            this.V_MESSAGE.AppInterfaceAdv = null;
            this.V_MESSAGE.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.V_MESSAGE.Location = new System.Drawing.Point(8, 163);
            this.V_MESSAGE.Name = "V_MESSAGE";
            this.V_MESSAGE.PromptAlignHoriz = InfoSummit.Win.ControlAdv.ISUtil.Enum.AlignHoriz.Center;
            this.V_MESSAGE.PromptText = "Set Message";
            isLanguageElement3.Default = "Set Message";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "Set Message";
            isLanguageElement3.TL2_CN = "";
            isLanguageElement3.TL3_VN = "";
            isLanguageElement3.TL4_JP = "";
            isLanguageElement3.TL5_XAA = "";
            this.V_MESSAGE.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.V_MESSAGE.Size = new System.Drawing.Size(686, 29);
            this.V_MESSAGE.TabIndex = 194;
            this.V_MESSAGE.TabStop = false;
            // 
            // V_PB_INTERFACE
            // 
            this.V_PB_INTERFACE.BarDividersCount = 2;
            this.V_PB_INTERFACE.BarFillPercent = 32F;
            this.V_PB_INTERFACE.Location = new System.Drawing.Point(98, 125);
            this.V_PB_INTERFACE.Name = "V_PB_INTERFACE";
            this.V_PB_INTERFACE.Size = new System.Drawing.Size(472, 32);
            this.V_PB_INTERFACE.StepSize = 10F;
            this.V_PB_INTERFACE.TabIndex = 193;
            this.V_PB_INTERFACE.Text = "SECOM";
            // 
            // V_START_ROW
            // 
            this.V_START_ROW.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.V_START_ROW.AutoScroll = true;
            this.V_START_ROW.ComboBoxValue = "";
            this.V_START_ROW.ComboData = null;
            this.V_START_ROW.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_START_ROW.DataAdapter = null;
            this.V_START_ROW.DataColumn = null;
            this.V_START_ROW.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.V_START_ROW.DoubleValue = 0D;
            this.V_START_ROW.EditAdvType = InfoSummit.Win.ControlAdv.ISUtil.Enum.EditAdvType.NumberEdit;
            this.V_START_ROW.EditValue = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.V_START_ROW.Insertable = false;
            this.V_START_ROW.Location = new System.Drawing.Point(8, 81);
            this.V_START_ROW.LookupAdapter = null;
            this.V_START_ROW.Name = "V_START_ROW";
            this.V_START_ROW.NumberValue = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.V_START_ROW.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_START_ROW.PromptText = "Start Row";
            isLanguageElement4.Default = "Start Row";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "Start Row";
            isLanguageElement4.TL2_CN = null;
            isLanguageElement4.TL3_VN = null;
            isLanguageElement4.TL4_JP = null;
            isLanguageElement4.TL5_XAA = null;
            this.V_START_ROW.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.V_START_ROW.PromptWidth = 120;
            this.V_START_ROW.Size = new System.Drawing.Size(180, 21);
            this.V_START_ROW.TabIndex = 192;
            this.V_START_ROW.TextValue = "";
            this.V_START_ROW.Updatable = false;
            // 
            // UPLOAD_FILE_PATH
            // 
            this.UPLOAD_FILE_PATH.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.UPLOAD_FILE_PATH.AutoScroll = true;
            this.UPLOAD_FILE_PATH.ComboBoxValue = "";
            this.UPLOAD_FILE_PATH.ComboData = null;
            this.UPLOAD_FILE_PATH.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.UPLOAD_FILE_PATH.DataAdapter = null;
            this.UPLOAD_FILE_PATH.DataColumn = null;
            this.UPLOAD_FILE_PATH.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.UPLOAD_FILE_PATH.DoubleValue = 0D;
            this.UPLOAD_FILE_PATH.EditValue = "";
            this.UPLOAD_FILE_PATH.Insertable = false;
            this.UPLOAD_FILE_PATH.Location = new System.Drawing.Point(8, 54);
            this.UPLOAD_FILE_PATH.LookupAdapter = null;
            this.UPLOAD_FILE_PATH.Name = "UPLOAD_FILE_PATH";
            this.UPLOAD_FILE_PATH.Nullable = true;
            this.UPLOAD_FILE_PATH.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.UPLOAD_FILE_PATH.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.UPLOAD_FILE_PATH.PromptText = "Upload File";
            isLanguageElement5.Default = "Upload File";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "업로드 파일";
            isLanguageElement5.TL2_CN = null;
            isLanguageElement5.TL3_VN = null;
            isLanguageElement5.TL4_JP = null;
            isLanguageElement5.TL5_XAA = null;
            this.UPLOAD_FILE_PATH.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.UPLOAD_FILE_PATH.PromptWidth = 120;
            this.UPLOAD_FILE_PATH.ReadOnly = true;
            this.UPLOAD_FILE_PATH.Size = new System.Drawing.Size(687, 21);
            this.UPLOAD_FILE_PATH.TabIndex = 189;
            this.UPLOAD_FILE_PATH.TabStop = false;
            this.UPLOAD_FILE_PATH.TextValue = "";
            this.UPLOAD_FILE_PATH.Updatable = false;
            // 
            // BTN_CLOSED
            // 
            this.BTN_CLOSED.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_CLOSED.ButtonText = "Cancel";
            isLanguageElement6.Default = "Cancel";
            isLanguageElement6.SiteName = null;
            isLanguageElement6.TL1_KR = "취소";
            isLanguageElement6.TL2_CN = null;
            isLanguageElement6.TL3_VN = null;
            isLanguageElement6.TL4_JP = null;
            isLanguageElement6.TL5_XAA = null;
            this.BTN_CLOSED.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement6});
            this.BTN_CLOSED.Location = new System.Drawing.Point(594, 81);
            this.BTN_CLOSED.Name = "BTN_CLOSED";
            this.BTN_CLOSED.Size = new System.Drawing.Size(100, 24);
            this.BTN_CLOSED.TabIndex = 190;
            this.BTN_CLOSED.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_CLOSED_ButtonClick);
            // 
            // BTN_SELECT_EXCEL_FILE
            // 
            this.BTN_SELECT_EXCEL_FILE.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_SELECT_EXCEL_FILE.ButtonText = "Select File";
            isLanguageElement7.Default = "Select File";
            isLanguageElement7.SiteName = null;
            isLanguageElement7.TL1_KR = "파일선택";
            isLanguageElement7.TL2_CN = null;
            isLanguageElement7.TL3_VN = null;
            isLanguageElement7.TL4_JP = null;
            isLanguageElement7.TL5_XAA = null;
            this.BTN_SELECT_EXCEL_FILE.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement7});
            this.BTN_SELECT_EXCEL_FILE.Location = new System.Drawing.Point(346, 81);
            this.BTN_SELECT_EXCEL_FILE.Name = "BTN_SELECT_EXCEL_FILE";
            this.BTN_SELECT_EXCEL_FILE.Size = new System.Drawing.Size(100, 24);
            this.BTN_SELECT_EXCEL_FILE.TabIndex = 190;
            this.BTN_SELECT_EXCEL_FILE.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_SELECT_EXCEL_FILE_ButtonClick);
            // 
            // BTN_FILE_UPLOAD
            // 
            this.BTN_FILE_UPLOAD.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_FILE_UPLOAD.ButtonText = "Execel Import";
            isLanguageElement8.Default = "Execel Import";
            isLanguageElement8.SiteName = null;
            isLanguageElement8.TL1_KR = "자료 업로드";
            isLanguageElement8.TL2_CN = null;
            isLanguageElement8.TL3_VN = null;
            isLanguageElement8.TL4_JP = null;
            isLanguageElement8.TL5_XAA = null;
            this.BTN_FILE_UPLOAD.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement8});
            this.BTN_FILE_UPLOAD.Location = new System.Drawing.Point(470, 81);
            this.BTN_FILE_UPLOAD.Name = "BTN_FILE_UPLOAD";
            this.BTN_FILE_UPLOAD.Size = new System.Drawing.Size(100, 24);
            this.BTN_FILE_UPLOAD.TabIndex = 191;
            this.BTN_FILE_UPLOAD.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_FILE_UPLOAD_ButtonClick);
            // 
            // isPrompt1
            // 
            this.isPrompt1.AppInterfaceAdv = null;
            this.isPrompt1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.isPrompt1.Location = new System.Drawing.Point(8, 90);
            this.isPrompt1.Name = "isPrompt1";
            this.isPrompt1.PromptStyle = InfoSummit.Win.ControlAdv.ISUtil.Enum.PromptStyle.UnderLine;
            this.isPrompt1.PromptText = "";
            isLanguageElement9.Default = "";
            isLanguageElement9.SiteName = null;
            isLanguageElement9.TL1_KR = null;
            isLanguageElement9.TL2_CN = null;
            isLanguageElement9.TL3_VN = null;
            isLanguageElement9.TL4_JP = null;
            isLanguageElement9.TL5_XAA = null;
            this.isPrompt1.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement9});
            this.isPrompt1.Size = new System.Drawing.Size(692, 27);
            this.isPrompt1.TabIndex = 195;
            this.isPrompt1.TabStop = false;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.RestoreDirectory = true;
            // 
            // IDC_IMPORT_EXCEL
            // 
            isOraParamElement1.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement1.MemberControl = null;
            isOraParamElement1.MemberValue = null;
            isOraParamElement1.OraDbTypeString = "VARCHAR2";
            isOraParamElement1.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement1.ParamName = "P_PERSON_NUM";
            isOraParamElement1.Size = 0;
            isOraParamElement1.SourceColumn = null;
            isOraParamElement2.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement2.MemberControl = null;
            isOraParamElement2.MemberValue = null;
            isOraParamElement2.OraDbTypeString = "VARCHAR2";
            isOraParamElement2.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement2.ParamName = "P_NAME";
            isOraParamElement2.Size = 0;
            isOraParamElement2.SourceColumn = null;
            isOraParamElement3.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement3.MemberControl = null;
            isOraParamElement3.MemberValue = null;
            isOraParamElement3.OraDbTypeString = "DATE";
            isOraParamElement3.OraType = System.Data.OracleClient.OracleType.DateTime;
            isOraParamElement3.ParamName = "P_TRANSFER_DATE";
            isOraParamElement3.Size = 0;
            isOraParamElement3.SourceColumn = null;
            isOraParamElement4.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement4.MemberControl = null;
            isOraParamElement4.MemberValue = null;
            isOraParamElement4.OraDbTypeString = "NUMBER";
            isOraParamElement4.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement4.ParamName = "P_TRANSFER_AMOUNT";
            isOraParamElement4.Size = 22;
            isOraParamElement4.SourceColumn = null;
            isOraParamElement5.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement5.MemberControl = null;
            isOraParamElement5.MemberValue = null;
            isOraParamElement5.OraDbTypeString = "VARCHAR2";
            isOraParamElement5.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement5.ParamName = "P_DESCRIPTION";
            isOraParamElement5.Size = 0;
            isOraParamElement5.SourceColumn = null;
            isOraParamElement6.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement6.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement6.MemberValue = "SOB_ID";
            isOraParamElement6.OraDbTypeString = "NUMBER";
            isOraParamElement6.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement6.ParamName = "P_SOB_ID";
            isOraParamElement6.Size = 22;
            isOraParamElement6.SourceColumn = null;
            isOraParamElement7.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement7.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement7.MemberValue = "ORG_ID";
            isOraParamElement7.OraDbTypeString = "NUMBER";
            isOraParamElement7.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement7.ParamName = "P_ORG_ID";
            isOraParamElement7.Size = 22;
            isOraParamElement7.SourceColumn = null;
            isOraParamElement8.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement8.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement8.MemberValue = "USER_ID";
            isOraParamElement8.OraDbTypeString = "NUMBER";
            isOraParamElement8.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement8.ParamName = "P_USER_ID";
            isOraParamElement8.Size = 22;
            isOraParamElement8.SourceColumn = null;
            isOraParamElement9.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement9.MemberControl = null;
            isOraParamElement9.MemberValue = null;
            isOraParamElement9.OraDbTypeString = "VARCHAR2";
            isOraParamElement9.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement9.ParamName = "O_STATUS";
            isOraParamElement9.Size = 0;
            isOraParamElement9.SourceColumn = null;
            isOraParamElement10.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement10.MemberControl = null;
            isOraParamElement10.MemberValue = null;
            isOraParamElement10.OraDbTypeString = "VARCHAR2";
            isOraParamElement10.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement10.ParamName = "O_MESSAGE";
            isOraParamElement10.Size = 0;
            isOraParamElement10.SourceColumn = null;
            this.IDC_IMPORT_EXCEL.CommandParamElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraParamElement[] {
            isOraParamElement1,
            isOraParamElement2,
            isOraParamElement3,
            isOraParamElement4,
            isOraParamElement5,
            isOraParamElement6,
            isOraParamElement7,
            isOraParamElement8,
            isOraParamElement9,
            isOraParamElement10});
            this.IDC_IMPORT_EXCEL.DataTransaction = null;
            this.IDC_IMPORT_EXCEL.OraConnection = this.isOraConnection1;
            this.IDC_IMPORT_EXCEL.OraOwner = "APPS";
            this.IDC_IMPORT_EXCEL.OraPackage = "HRR_RETIRE_BANK_TRANSFER_G";
            this.IDC_IMPORT_EXCEL.OraProcedure = "IMPORT_EXCEL";
            // 
            // IDC_GET_LOCAL_DATETIME_P
            // 
            isOraParamElement11.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement11.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement11.MemberValue = "SOB_ID";
            isOraParamElement11.OraDbTypeString = "NUMBER";
            isOraParamElement11.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement11.ParamName = "P_SOB_ID";
            isOraParamElement11.Size = 22;
            isOraParamElement11.SourceColumn = null;
            isOraParamElement12.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement12.MemberControl = null;
            isOraParamElement12.MemberValue = null;
            isOraParamElement12.OraDbTypeString = "DATE";
            isOraParamElement12.OraType = System.Data.OracleClient.OracleType.DateTime;
            isOraParamElement12.ParamName = "X_LOCAL_DATE";
            isOraParamElement12.Size = 0;
            isOraParamElement12.SourceColumn = null;
            this.IDC_GET_LOCAL_DATETIME_P.CommandParamElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraParamElement[] {
            isOraParamElement11,
            isOraParamElement12});
            this.IDC_GET_LOCAL_DATETIME_P.DataTransaction = null;
            this.IDC_GET_LOCAL_DATETIME_P.OraConnection = this.isOraConnection1;
            this.IDC_GET_LOCAL_DATETIME_P.OraOwner = "APPS";
            this.IDC_GET_LOCAL_DATETIME_P.OraPackage = "EAPP_COMMON_G";
            this.IDC_GET_LOCAL_DATETIME_P.OraProcedure = "GET_LOCAL_DATETIME_P";
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GB_UPLOAD_FILE)).EndInit();
            this.GB_UPLOAD_FILE.ResumeLayout(false);

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISGroupBox GB_UPLOAD_FILE;
        private InfoSummit.Win.ControlAdv.ISEditAdv UPLOAD_FILE_PATH;
        private InfoSummit.Win.ControlAdv.ISButton BTN_FILE_UPLOAD;
        private InfoSummit.Win.ControlAdv.ISButton BTN_SELECT_EXCEL_FILE;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private InfoSummit.Win.ControlAdv.ISButton BTN_CLOSED;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_START_ROW;
        private InfoSummit.Win.ControlAdv.ISPrompt V_MESSAGE;
        private InfoSummit.Win.ControlAdv.ISProgressBar V_PB_INTERFACE;
        private InfoSummit.Win.ControlAdv.ISPrompt isPrompt1;
        private InfoSummit.Win.ControlAdv.ISDataCommand IDC_IMPORT_EXCEL;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_CORP_ID;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_STD_DATE;
        private InfoSummit.Win.ControlAdv.ISDataCommand IDC_GET_LOCAL_DATETIME_P;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_SYSDATE;
    }
}