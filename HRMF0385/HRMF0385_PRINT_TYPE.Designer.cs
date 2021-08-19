namespace HRMF0385
{
    partial class HRMF0385_PRINT_TYPE
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
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement8 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement6 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement7 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isMessageAdapter1 = new InfoSummit.Win.ControlAdv.ISMessageAdapter(this.components);
            this.BTN_OK = new InfoSummit.Win.ControlAdv.ISButton();
            this.BTN_CLOSED = new InfoSummit.Win.ControlAdv.ISButton();
            this.RB_PDF = new InfoSummit.Win.ControlAdv.ISRadioButtonAdv();
            this.isGroupBox7 = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.isRadioButtonAdv2 = new InfoSummit.Win.ControlAdv.ISRadioButtonAdv();
            this.isRadioButtonAdv3 = new InfoSummit.Win.ControlAdv.ISRadioButtonAdv();
            this.V_PRINT_TYPE = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.isRadioButtonAdv1 = new InfoSummit.Win.ControlAdv.ISRadioButtonAdv();
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RB_PDF)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox7)).BeginInit();
            this.isGroupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.isRadioButtonAdv2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.isRadioButtonAdv3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.isRadioButtonAdv1)).BeginInit();
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
            // isMessageAdapter1
            // 
            this.isMessageAdapter1.OraConnection = this.isOraConnection1;
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
            // HRMF0385_PRINT_TYPE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(259, 124);
            this.Controls.Add(this.isGroupBox7);
            this.Controls.Add(this.BTN_CLOSED);
            this.Controls.Add(this.BTN_OK);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "HRMF0385_PRINT_TYPE";
            this.Padding = new System.Windows.Forms.Padding(2);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Printer Type";
            this.Load += new System.EventHandler(this.HRMF0385_PRINT_TYPE_Load);
            this.BTN_OK.Location = new System.Drawing.Point(72, 87);
            this.BTN_OK.Name = "BTN_OK";
            this.BTN_OK.Size = new System.Drawing.Size(83, 25);
            this.BTN_OK.TabIndex = 1;
            this.BTN_OK.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.BTN_OK.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_OK_ButtonClick);
            // 
            // BTN_CLOSED
            // 
            this.BTN_CLOSED.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_CLOSED.ButtonText = "닫기";
            isLanguageElement8.Default = "Closed";
            isLanguageElement8.SiteName = null;
            isLanguageElement8.TL1_KR = "닫기";
            isLanguageElement8.TL2_CN = null;
            isLanguageElement8.TL3_VN = null;
            isLanguageElement8.TL4_JP = null;
            isLanguageElement8.TL5_XAA = null;
            this.BTN_CLOSED.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement8});
            this.BTN_CLOSED.Location = new System.Drawing.Point(164, 87);
            this.BTN_CLOSED.Name = "BTN_CLOSED";
            this.BTN_CLOSED.Size = new System.Drawing.Size(83, 25);
            this.BTN_CLOSED.TabIndex = 2;
            this.BTN_CLOSED.TerritoryLanguage = InfoSummit.Win.ControlAdv.ISUtil.Enum.TerritoryLanguage.TL1_KR;
            this.BTN_CLOSED.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_CLOSED_ButtonClick);
            // 
            // RB_PDF
            // 
            this.RB_PDF.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.RB_PDF.CheckedString = "PDF";
            this.RB_PDF.DataAdapter = null;
            this.RB_PDF.DataColumn = null;
            this.RB_PDF.Location = new System.Drawing.Point(129, 10);
            this.RB_PDF.MetroColor = System.Drawing.Color.Empty;
            this.RB_PDF.Name = "RB_PDF";
            this.RB_PDF.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.RB_PDF.PromptText = "Preview";
            isLanguageElement6.Default = "Preview";
            isLanguageElement6.SiteName = null;
            isLanguageElement6.TL1_KR = "미리보기";
            isLanguageElement6.TL2_CN = null;
            isLanguageElement6.TL3_VN = null;
            isLanguageElement6.TL4_JP = null;
            isLanguageElement6.TL5_XAA = null;
            this.RB_PDF.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement6});
            this.RB_PDF.RadioButtonValue = null;
            this.RB_PDF.RadioCheckedString = "PREVIEW";
            this.RB_PDF.Size = new System.Drawing.Size(113, 21);
            this.RB_PDF.Style = Syncfusion.Windows.Forms.Tools.RadioButtonAdvStyle.Office2007;
            this.RB_PDF.TabIndex = 2;
            this.RB_PDF.TabStop = false;
            this.RB_PDF.Text = "Preview";
            this.RB_PDF.ThemesEnabled = false;
            this.RB_PDF.CheckChanged += new System.EventHandler(this.isRadioButtonAdv_CheckChanged);
            // 
            // isGroupBox7
            // 
            this.isGroupBox7.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isGroupBox7.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.isGroupBox7.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.isGroupBox7.Controls.Add(this.isRadioButtonAdv2);
            this.isGroupBox7.Controls.Add(this.isRadioButtonAdv3);
            this.isGroupBox7.Controls.Add(this.V_PRINT_TYPE);
            this.isGroupBox7.Controls.Add(this.isRadioButtonAdv1);
            this.isGroupBox7.Controls.Add(this.RB_PDF);
            this.isGroupBox7.Location = new System.Drawing.Point(5, 5);
            this.isGroupBox7.Name = "isGroupBox7";
            this.isGroupBox7.PromptText = "isGroupBox4";
            isLanguageElement7.Default = "isGroupBox4";
            isLanguageElement7.SiteName = null;
            isLanguageElement7.TL1_KR = null;
            isLanguageElement7.TL2_CN = null;
            isLanguageElement7.TL3_VN = null;
            isLanguageElement7.TL4_JP = null;
            isLanguageElement7.TL5_XAA = null;
            this.isGroupBox7.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement7});
            this.isGroupBox7.PromptVisible = false;
            this.isGroupBox7.Size = new System.Drawing.Size(249, 69);
            this.isGroupBox7.TabIndex = 0;
            // 
            // isRadioButtonAdv2
            // 
            this.isRadioButtonAdv2.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isRadioButtonAdv2.CheckedString = "PRINT";
            this.isRadioButtonAdv2.DataAdapter = null;
            this.isRadioButtonAdv2.DataColumn = null;
            this.isRadioButtonAdv2.Location = new System.Drawing.Point(10, 37);
            this.isRadioButtonAdv2.MetroColor = System.Drawing.Color.Empty;
            this.isRadioButtonAdv2.Name = "isRadioButtonAdv2";
            this.isRadioButtonAdv2.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.isRadioButtonAdv2.PromptText = "Excel Export";
            isLanguageElement2.Default = "Excel Export";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "엑셀 저장";
            isLanguageElement2.TL2_CN = null;
            isLanguageElement2.TL3_VN = null;
            isLanguageElement2.TL4_JP = null;
            isLanguageElement2.TL5_XAA = null;
            this.isRadioButtonAdv2.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.isRadioButtonAdv2.RadioButtonValue = null;
            this.isRadioButtonAdv2.RadioCheckedString = "EXCEL";
            this.isRadioButtonAdv2.Size = new System.Drawing.Size(113, 21);
            this.isRadioButtonAdv2.Style = Syncfusion.Windows.Forms.Tools.RadioButtonAdvStyle.Office2007;
            this.isRadioButtonAdv2.TabIndex = 440;
            this.isRadioButtonAdv2.Text = "Excel Export";
            this.isRadioButtonAdv2.ThemesEnabled = false;
            this.isRadioButtonAdv2.CheckChanged += new System.EventHandler(this.isRadioButtonAdv_CheckChanged);
            // 
            // isRadioButtonAdv3
            // 
            this.isRadioButtonAdv3.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isRadioButtonAdv3.CheckedString = "PDF";
            this.isRadioButtonAdv3.DataAdapter = null;
            this.isRadioButtonAdv3.DataColumn = null;
            this.isRadioButtonAdv3.Location = new System.Drawing.Point(129, 37);
            this.isRadioButtonAdv3.MetroColor = System.Drawing.Color.Empty;
            this.isRadioButtonAdv3.Name = "isRadioButtonAdv3";
            this.isRadioButtonAdv3.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.isRadioButtonAdv3.PromptText = "PDF";
            isLanguageElement3.Default = "PDF";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "PDF";
            isLanguageElement3.TL2_CN = null;
            isLanguageElement3.TL3_VN = null;
            isLanguageElement3.TL4_JP = null;
            isLanguageElement3.TL5_XAA = null;
            this.isRadioButtonAdv3.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.isRadioButtonAdv3.RadioButtonValue = null;
            this.isRadioButtonAdv3.RadioCheckedString = "PDF";
            this.isRadioButtonAdv3.Size = new System.Drawing.Size(113, 21);
            this.isRadioButtonAdv3.Style = Syncfusion.Windows.Forms.Tools.RadioButtonAdvStyle.Office2007;
            this.isRadioButtonAdv3.TabIndex = 441;
            this.isRadioButtonAdv3.TabStop = false;
            this.isRadioButtonAdv3.Text = "PDF";
            this.isRadioButtonAdv3.ThemesEnabled = false;
            this.isRadioButtonAdv3.CheckChanged += new System.EventHandler(this.isRadioButtonAdv_CheckChanged);
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
            this.V_PRINT_TYPE.Location = new System.Drawing.Point(195, 9);
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
            isLanguageElement4.Default = "Print Type";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "인쇄 구분";
            isLanguageElement4.TL2_CN = "";
            isLanguageElement4.TL3_VN = "";
            isLanguageElement4.TL4_JP = "";
            isLanguageElement4.TL5_XAA = "";
            this.V_PRINT_TYPE.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.V_PRINT_TYPE.PromptVisible = false;
            this.V_PRINT_TYPE.PromptWidth = 80;
            this.V_PRINT_TYPE.ReadOnly = true;
            this.V_PRINT_TYPE.Size = new System.Drawing.Size(27, 21);
            this.V_PRINT_TYPE.TabIndex = 439;
            this.V_PRINT_TYPE.TextValue = "";
            this.V_PRINT_TYPE.Visible = false;
            // 
            // isRadioButtonAdv1
            // 
            this.isRadioButtonAdv1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isRadioButtonAdv1.CheckedString = "PRINT";
            this.isRadioButtonAdv1.DataAdapter = null;
            this.isRadioButtonAdv1.DataColumn = null;
            this.isRadioButtonAdv1.Location = new System.Drawing.Point(10, 10);
            this.isRadioButtonAdv1.MetroColor = System.Drawing.Color.Empty;
            this.isRadioButtonAdv1.Name = "isRadioButtonAdv1";
            this.isRadioButtonAdv1.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.isRadioButtonAdv1.PromptText = "Printer";
            isLanguageElement5.Default = "Printer";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "프린터 인쇄";
            isLanguageElement5.TL2_CN = null;
            isLanguageElement5.TL3_VN = null;
            isLanguageElement5.TL4_JP = null;
            isLanguageElement5.TL5_XAA = null;
            this.isRadioButtonAdv1.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.isRadioButtonAdv1.RadioButtonValue = null;
            this.isRadioButtonAdv1.RadioCheckedString = "PRINT";
            this.isRadioButtonAdv1.Size = new System.Drawing.Size(113, 21);
            this.isRadioButtonAdv1.Style = Syncfusion.Windows.Forms.Tools.RadioButtonAdvStyle.Office2007;
            this.isRadioButtonAdv1.TabIndex = 1;
            this.isRadioButtonAdv1.Text = "Printer";
            this.isRadioButtonAdv1.ThemesEnabled = false;
            this.isRadioButtonAdv1.CheckChanged += new System.EventHandler(this.isRadioButtonAdv_CheckChanged);
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.RB_PDF)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox7)).EndInit();
            this.isGroupBox7.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.isRadioButtonAdv2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.isRadioButtonAdv3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.isRadioButtonAdv1)).EndInit();

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter isMessageAdapter1;
        private InfoSummit.Win.ControlAdv.ISButton BTN_CLOSED;
        private InfoSummit.Win.ControlAdv.ISButton BTN_OK;
        private InfoSummit.Win.ControlAdv.ISRadioButtonAdv RB_PDF;
        private InfoSummit.Win.ControlAdv.ISGroupBox isGroupBox7;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_PRINT_TYPE;
        private InfoSummit.Win.ControlAdv.ISRadioButtonAdv isRadioButtonAdv1;
        private InfoSummit.Win.ControlAdv.ISRadioButtonAdv isRadioButtonAdv2;
        private InfoSummit.Win.ControlAdv.ISRadioButtonAdv isRadioButtonAdv3;
    }
}

