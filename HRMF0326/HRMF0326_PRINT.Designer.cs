namespace HRMF0326
{
    partial class HRMF0326_PRINT
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
            InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo oraConnectionInfo1 = new InfoSummit.Win.ControlAdv.ISDataUtil.OraConnectionInfo();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement1 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement7 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement6 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.BTN_OK = new InfoSummit.Win.ControlAdv.ISButton();
            this.BTN_CLOSED = new InfoSummit.Win.ControlAdv.ISButton();
            this.RB_PRINTER = new InfoSummit.Win.ControlAdv.ISRadioButtonAdv();
            this.RB_PDF = new InfoSummit.Win.ControlAdv.ISRadioButtonAdv();
            this.isGroupBox7 = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.V_PRINT_TYPE = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.CB_PRINT_PREVIEW = new InfoSummit.Win.ControlAdv.ISCheckBoxAdv();
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RB_PRINTER)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RB_PDF)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox7)).BeginInit();
            this.isGroupBox7.SuspendLayout();
            // 
            // isAppInterfaceAdv1
            // 
            this.isAppInterfaceAdv1.AppMainButtonClick += new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv.ButtonEventHandler(this.isAppInterfaceAdv1_AppMainButtonClick);
            // 
            // isOraConnection1
            // 
            this.isOraConnection1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isOraConnection1.OraConnectionInfo = oraConnectionInfo1;
            this.isOraConnection1.OraHost = "211.168.59.26";
            this.isOraConnection1.OraPassword = "infoflex";
            this.isOraConnection1.OraPort = "1521";
            this.isOraConnection1.OraServiceName = "fxcdb";
            this.isOraConnection1.OraUserId = "APPS";
            // 
            // BTN_OK
            // 
            this.BTN_OK.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_OK.ButtonText = "확인";
            isLanguageElement1.Default = "O.K";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = "확인";
            isLanguageElement1.TL2_CN = null;
            isLanguageElement1.TL3_VN = null;
            isLanguageElement1.TL4_JP = null;
            isLanguageElement1.TL5_XAA = null;
            this.BTN_OK.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            // 
            // HRMF0326_PRINT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(238, 112);
            this.ControlBox = false;
            this.Controls.Add(this.V_PRINT_TYPE);
            this.Controls.Add(this.isGroupBox7);
            this.Controls.Add(this.BTN_CLOSED);
            this.Controls.Add(this.BTN_OK);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "HRMF0326_PRINT";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "◈ 인쇄 확인";
            this.Load += new System.EventHandler(this.HRMF0326_PRINT_Load);
            this.BTN_OK.Location = new System.Drawing.Point(55, 77);
            this.BTN_OK.Name = "BTN_OK";
            this.BTN_OK.Size = new System.Drawing.Size(83, 25);
            this.BTN_OK.TabIndex = 12;
            this.BTN_OK.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.BTN_OK.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_OK_ButtonClick);
            // 
            // BTN_CLOSED
            // 
            this.BTN_CLOSED.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_CLOSED.ButtonText = "닫기";
            isLanguageElement7.Default = "Closed";
            isLanguageElement7.SiteName = null;
            isLanguageElement7.TL1_KR = "닫기";
            isLanguageElement7.TL2_CN = null;
            isLanguageElement7.TL3_VN = null;
            isLanguageElement7.TL4_JP = null;
            isLanguageElement7.TL5_XAA = null;
            this.BTN_CLOSED.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement7});
            this.BTN_CLOSED.Location = new System.Drawing.Point(147, 77);
            this.BTN_CLOSED.Name = "BTN_CLOSED";
            this.BTN_CLOSED.Size = new System.Drawing.Size(83, 25);
            this.BTN_CLOSED.TabIndex = 13;
            this.BTN_CLOSED.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.BTN_CLOSED.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_CLOSED_ButtonClick);
            // 
            // RB_PRINTER
            // 
            this.RB_PRINTER.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.RB_PRINTER.CheckedString = "PRINT";
            this.RB_PRINTER.DataAdapter = null;
            this.RB_PRINTER.DataColumn = null;
            this.RB_PRINTER.Location = new System.Drawing.Point(6, 9);
            this.RB_PRINTER.MetroColor = System.Drawing.Color.Empty;
            this.RB_PRINTER.Name = "RB_PRINTER";
            this.RB_PRINTER.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.RB_PRINTER.PromptText = "Printer";
            isLanguageElement4.Default = "Printer";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "Printer";
            isLanguageElement4.TL2_CN = null;
            isLanguageElement4.TL3_VN = null;
            isLanguageElement4.TL4_JP = null;
            isLanguageElement4.TL5_XAA = null;
            this.RB_PRINTER.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.RB_PRINTER.RadioButtonValue = null;
            this.RB_PRINTER.RadioCheckedString = "PRINTER";
            this.RB_PRINTER.Size = new System.Drawing.Size(124, 21);
            this.RB_PRINTER.Style = Syncfusion.Windows.Forms.Tools.RadioButtonAdvStyle.Office2007;
            this.RB_PRINTER.TabIndex = 436;
            this.RB_PRINTER.Text = "Printer";
            this.RB_PRINTER.ThemesEnabled = false;
            this.RB_PRINTER.CheckChanged += new System.EventHandler(this.isRadioButtonAdv_CheckChanged);
            // 
            // RB_PDF
            // 
            this.RB_PDF.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.RB_PDF.CheckedString = "PDF";
            this.RB_PDF.DataAdapter = null;
            this.RB_PDF.DataColumn = null;
            this.RB_PDF.Location = new System.Drawing.Point(139, 9);
            this.RB_PDF.MetroColor = System.Drawing.Color.Empty;
            this.RB_PDF.Name = "RB_PDF";
            this.RB_PDF.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.RB_PDF.PromptText = "PDF";
            isLanguageElement5.Default = "PDF";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "PDF";
            isLanguageElement5.TL2_CN = null;
            isLanguageElement5.TL3_VN = null;
            isLanguageElement5.TL4_JP = null;
            isLanguageElement5.TL5_XAA = null;
            this.RB_PDF.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.RB_PDF.RadioButtonValue = null;
            this.RB_PDF.RadioCheckedString = "PDF";
            this.RB_PDF.Size = new System.Drawing.Size(81, 21);
            this.RB_PDF.Style = Syncfusion.Windows.Forms.Tools.RadioButtonAdvStyle.Office2007;
            this.RB_PDF.TabIndex = 437;
            this.RB_PDF.Text = "PDF";
            this.RB_PDF.ThemesEnabled = false;
            this.RB_PDF.CheckChanged += new System.EventHandler(this.isRadioButtonAdv_CheckChanged);
            // 
            // isGroupBox7
            // 
            this.isGroupBox7.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isGroupBox7.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.isGroupBox7.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.isGroupBox7.Controls.Add(this.CB_PRINT_PREVIEW);
            this.isGroupBox7.Controls.Add(this.RB_PRINTER);
            this.isGroupBox7.Controls.Add(this.RB_PDF);
            this.isGroupBox7.Location = new System.Drawing.Point(8, 8);
            this.isGroupBox7.Name = "isGroupBox7";
            this.isGroupBox7.PromptText = "isGroupBox4";
            isLanguageElement6.Default = "isGroupBox4";
            isLanguageElement6.SiteName = null;
            isLanguageElement6.TL1_KR = null;
            isLanguageElement6.TL2_CN = null;
            isLanguageElement6.TL3_VN = null;
            isLanguageElement6.TL4_JP = null;
            isLanguageElement6.TL5_XAA = null;
            this.isGroupBox7.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement6});
            this.isGroupBox7.PromptVisible = false;
            this.isGroupBox7.Size = new System.Drawing.Size(223, 60);
            this.isGroupBox7.TabIndex = 438;
            // 
            // V_PRINT_TYPE
            // 
            this.V_PRINT_TYPE.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.V_PRINT_TYPE.ComboBoxValue = "";
            this.V_PRINT_TYPE.ComboData = null;
            this.V_PRINT_TYPE.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_PRINT_TYPE.DataAdapter = null;
            this.V_PRINT_TYPE.DataColumn = null;
            this.V_PRINT_TYPE.DateTimeValue = new System.DateTime(2010, 3, 17, 19, 7, 59, 703);
            this.V_PRINT_TYPE.DoubleValue = 0D;
            this.V_PRINT_TYPE.EditValue = "";
            this.V_PRINT_TYPE.Location = new System.Drawing.Point(8, 82);
            this.V_PRINT_TYPE.LookupAdapter = null;
            this.V_PRINT_TYPE.Name = "V_PRINT_TYPE";
            this.V_PRINT_TYPE.Nullable = true;
            this.V_PRINT_TYPE.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_PRINT_TYPE.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_PRINT_TYPE.PromptText = "Print Type";
            isLanguageElement2.Default = "Print Type";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "인쇄 구분";
            isLanguageElement2.TL2_CN = "";
            isLanguageElement2.TL3_VN = "";
            isLanguageElement2.TL4_JP = "";
            isLanguageElement2.TL5_XAA = "";
            this.V_PRINT_TYPE.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.V_PRINT_TYPE.PromptVisible = false;
            this.V_PRINT_TYPE.PromptWidth = 80;
            this.V_PRINT_TYPE.ReadOnly = true;
            this.V_PRINT_TYPE.Size = new System.Drawing.Size(27, 21);
            this.V_PRINT_TYPE.TabIndex = 439;
            this.V_PRINT_TYPE.TextValue = "";
            this.V_PRINT_TYPE.Visible = false;
            // 
            // CB_PRINT_PREVIEW
            // 
            this.CB_PRINT_PREVIEW.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.CB_PRINT_PREVIEW.CheckBoxValue = "N";
            this.CB_PRINT_PREVIEW.CheckedString = "Y";
            this.CB_PRINT_PREVIEW.DataAdapter = null;
            this.CB_PRINT_PREVIEW.DataColumn = null;
            this.CB_PRINT_PREVIEW.Location = new System.Drawing.Point(6, 33);
            this.CB_PRINT_PREVIEW.Name = "CB_PRINT_PREVIEW";
            this.CB_PRINT_PREVIEW.PromptText = "Print Preview";
            isLanguageElement3.Default = "Print Preview";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "인쇄 미리보기";
            isLanguageElement3.TL2_CN = null;
            isLanguageElement3.TL3_VN = null;
            isLanguageElement3.TL4_JP = null;
            isLanguageElement3.TL5_XAA = null;
            this.CB_PRINT_PREVIEW.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.CB_PRINT_PREVIEW.Size = new System.Drawing.Size(124, 21);
            this.CB_PRINT_PREVIEW.TabIndex = 438;
            this.CB_PRINT_PREVIEW.UncheckedString = "N";
            // 
            // isMessageAdapter1
            // 
            this.isMessageAdapter1.OraConnection = this.isOraConnection1;
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.RB_PRINTER)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RB_PDF)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox7)).EndInit();
            this.isGroupBox7.ResumeLayout(false);

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISButton BTN_CLOSED;
        private InfoSummit.Win.ControlAdv.ISButton BTN_OK;
        private InfoSummit.Win.ControlAdv.ISRadioButtonAdv RB_PDF;
        private InfoSummit.Win.ControlAdv.ISRadioButtonAdv RB_PRINTER;
        private InfoSummit.Win.ControlAdv.ISGroupBox isGroupBox7;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_PRINT_TYPE;
        private InfoSummit.Win.ControlAdv.ISCheckBoxAdv CB_PRINT_PREVIEW;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
    }
}

