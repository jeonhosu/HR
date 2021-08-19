namespace HRMF0761
{
    partial class NTS_Reader
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
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement9 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement8 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement1 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement2 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement3 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement4 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement5 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement6 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement7 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            this.grpUtf8 = new System.Windows.Forms.GroupBox();
            this.txtUtf8 = new System.Windows.Forms.RichTextBox();
            this.V_FILE_NAME = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isGroupBox1 = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.V_PERSON_NUM = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.V_NAME = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.V_YYYYMM = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.BTN_CANCEL = new InfoSummit.Win.ControlAdv.ISButton();
            this.BTN_PDF_IMPORT = new InfoSummit.Win.ControlAdv.ISButton();
            this.BTN_FILE_FIND = new InfoSummit.Win.ControlAdv.ISButton();
            this.V_FILE_PWD = new InfoSummit.Win.ControlAdv.ISEditAdv();
            this.grpUtf8.SuspendLayout();
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox1)).BeginInit();
            this.isGroupBox1.SuspendLayout();
            // 
            // grpUtf8
            // 
            this.grpUtf8.Controls.Add(this.txtUtf8);
            this.grpUtf8.Location = new System.Drawing.Point(12, 151);
            this.grpUtf8.Name = "grpUtf8";
            this.grpUtf8.Size = new System.Drawing.Size(702, 199);
            this.grpUtf8.TabIndex = 1;
            this.grpUtf8.TabStop = false;
            this.grpUtf8.Text = "XML (UTF-8)";
            this.grpUtf8.Visible = false;
            // 
            // txtUtf8
            // 
            this.txtUtf8.BackColor = System.Drawing.SystemColors.Window;
            this.txtUtf8.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtUtf8.Font = new System.Drawing.Font("굴림", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.txtUtf8.Location = new System.Drawing.Point(6, 20);
            this.txtUtf8.Name = "txtUtf8";
            this.txtUtf8.ReadOnly = true;
            this.txtUtf8.Size = new System.Drawing.Size(690, 173);
            this.txtUtf8.TabIndex = 0;
            this.txtUtf8.TabStop = false;
            this.txtUtf8.Text = "";
            this.txtUtf8.Visible = false;
            // 
            // V_FILE_NAME
            // 
            this.V_FILE_NAME.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.V_FILE_NAME.ComboBoxValue = "";
            this.V_FILE_NAME.ComboData = null;
            this.V_FILE_NAME.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_FILE_NAME.DataAdapter = null;
            this.V_FILE_NAME.DataColumn = null;
            this.V_FILE_NAME.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.V_FILE_NAME.DoubleValue = 0;
            this.V_FILE_NAME.EditValue = "";
            // 
            // NTS_Reader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(726, 126);
            this.ControlBox = false;
            this.Controls.Add(this.isGroupBox1);
            this.Controls.Add(this.grpUtf8);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "NTS_Reader";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Pdf Reader";
            this.Load += new System.EventHandler(this.NTS_Reader_Load);
            this.V_FILE_NAME.Location = new System.Drawing.Point(7, 42);
            this.V_FILE_NAME.LookupAdapter = null;
            this.V_FILE_NAME.Name = "V_FILE_NAME";
            this.V_FILE_NAME.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_FILE_NAME.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_FILE_NAME.PromptText = "Pdf File";
            isLanguageElement9.Default = "Pdf File";
            isLanguageElement9.SiteName = null;
            isLanguageElement9.TL1_KR = "Pdf 파일";
            isLanguageElement9.TL2_CN = null;
            isLanguageElement9.TL3_VN = null;
            isLanguageElement9.TL4_JP = null;
            isLanguageElement9.TL5_XAA = null;
            this.V_FILE_NAME.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement9});
            this.V_FILE_NAME.Size = new System.Drawing.Size(525, 21);
            this.V_FILE_NAME.TabIndex = 5;
            this.V_FILE_NAME.TextValue = "";
            // 
            // isGroupBox1
            // 
            this.isGroupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.isGroupBox1.AppInterfaceAdv = null;
            this.isGroupBox1.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.isGroupBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.isGroupBox1.Controls.Add(this.V_PERSON_NUM);
            this.isGroupBox1.Controls.Add(this.V_NAME);
            this.isGroupBox1.Controls.Add(this.V_YYYYMM);
            this.isGroupBox1.Controls.Add(this.BTN_CANCEL);
            this.isGroupBox1.Controls.Add(this.BTN_PDF_IMPORT);
            this.isGroupBox1.Controls.Add(this.V_FILE_NAME);
            this.isGroupBox1.Controls.Add(this.BTN_FILE_FIND);
            this.isGroupBox1.Controls.Add(this.V_FILE_PWD);
            this.isGroupBox1.Location = new System.Drawing.Point(12, 12);
            this.isGroupBox1.Name = "isGroupBox1";
            this.isGroupBox1.PromptText = "isGroupBox1";
            isLanguageElement8.Default = "isGroupBox1";
            isLanguageElement8.SiteName = null;
            isLanguageElement8.TL1_KR = null;
            isLanguageElement8.TL2_CN = null;
            isLanguageElement8.TL3_VN = null;
            isLanguageElement8.TL4_JP = null;
            isLanguageElement8.TL5_XAA = null;
            this.isGroupBox1.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement8});
            this.isGroupBox1.PromptVisible = false;
            this.isGroupBox1.Size = new System.Drawing.Size(702, 101);
            this.isGroupBox1.TabIndex = 2;
            // 
            // V_PERSON_NUM
            // 
            this.V_PERSON_NUM.AppInterfaceAdv = null;
            this.V_PERSON_NUM.ComboBoxValue = "";
            this.V_PERSON_NUM.ComboData = null;
            this.V_PERSON_NUM.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_PERSON_NUM.DataAdapter = null;
            this.V_PERSON_NUM.DataColumn = null;
            this.V_PERSON_NUM.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.V_PERSON_NUM.DoubleValue = 0;
            this.V_PERSON_NUM.EditValue = "";
            this.V_PERSON_NUM.Location = new System.Drawing.Point(425, 14);
            this.V_PERSON_NUM.LookupAdapter = null;
            this.V_PERSON_NUM.Name = "V_PERSON_NUM";
            this.V_PERSON_NUM.Nullable = true;
            this.V_PERSON_NUM.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_PERSON_NUM.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_PERSON_NUM.PromptText = "Name";
            isLanguageElement1.Default = "Name";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = "성명";
            isLanguageElement1.TL2_CN = null;
            isLanguageElement1.TL3_VN = null;
            isLanguageElement1.TL4_JP = null;
            isLanguageElement1.TL5_XAA = null;
            this.V_PERSON_NUM.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            this.V_PERSON_NUM.PromptVisible = false;
            this.V_PERSON_NUM.ReadOnly = true;
            this.V_PERSON_NUM.Size = new System.Drawing.Size(107, 21);
            this.V_PERSON_NUM.TabIndex = 195;
            this.V_PERSON_NUM.TabStop = false;
            this.V_PERSON_NUM.TextValue = "";
            // 
            // V_NAME
            // 
            this.V_NAME.AppInterfaceAdv = null;
            this.V_NAME.ComboBoxValue = "";
            this.V_NAME.ComboData = null;
            this.V_NAME.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_NAME.DataAdapter = null;
            this.V_NAME.DataColumn = null;
            this.V_NAME.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.V_NAME.DoubleValue = 0;
            this.V_NAME.EditValue = "";
            this.V_NAME.Location = new System.Drawing.Point(219, 14);
            this.V_NAME.LookupAdapter = null;
            this.V_NAME.Name = "V_NAME";
            this.V_NAME.Nullable = true;
            this.V_NAME.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_NAME.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_NAME.PromptText = "Name";
            isLanguageElement2.Default = "Name";
            isLanguageElement2.SiteName = null;
            isLanguageElement2.TL1_KR = "성명";
            isLanguageElement2.TL2_CN = null;
            isLanguageElement2.TL3_VN = null;
            isLanguageElement2.TL4_JP = null;
            isLanguageElement2.TL5_XAA = null;
            this.V_NAME.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement2});
            this.V_NAME.ReadOnly = true;
            this.V_NAME.Size = new System.Drawing.Size(206, 21);
            this.V_NAME.TabIndex = 194;
            this.V_NAME.TabStop = false;
            this.V_NAME.TextValue = "";
            // 
            // V_YYYYMM
            // 
            this.V_YYYYMM.AppInterfaceAdv = null;
            this.V_YYYYMM.ComboBoxValue = "";
            this.V_YYYYMM.ComboData = null;
            this.V_YYYYMM.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_YYYYMM.DataAdapter = null;
            this.V_YYYYMM.DataColumn = null;
            this.V_YYYYMM.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.V_YYYYMM.DoubleValue = 0;
            this.V_YYYYMM.EditValue = "";
            this.V_YYYYMM.Location = new System.Drawing.Point(7, 14);
            this.V_YYYYMM.LookupAdapter = null;
            this.V_YYYYMM.Name = "V_YYYYMM";
            this.V_YYYYMM.Nullable = true;
            this.V_YYYYMM.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_YYYYMM.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_YYYYMM.PromptText = "Period Name";
            isLanguageElement3.Default = "Period Name";
            isLanguageElement3.SiteName = null;
            isLanguageElement3.TL1_KR = "정산년월";
            isLanguageElement3.TL2_CN = null;
            isLanguageElement3.TL3_VN = null;
            isLanguageElement3.TL4_JP = null;
            isLanguageElement3.TL5_XAA = null;
            this.V_YYYYMM.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement3});
            this.V_YYYYMM.ReadOnly = true;
            this.V_YYYYMM.Size = new System.Drawing.Size(206, 21);
            this.V_YYYYMM.TabIndex = 193;
            this.V_YYYYMM.TabStop = false;
            this.V_YYYYMM.TextValue = "";
            // 
            // BTN_CANCEL
            // 
            this.BTN_CANCEL.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_CANCEL.ButtonText = "Cancel";
            isLanguageElement4.Default = "Cancel";
            isLanguageElement4.SiteName = null;
            isLanguageElement4.TL1_KR = "취소";
            isLanguageElement4.TL2_CN = null;
            isLanguageElement4.TL3_VN = null;
            isLanguageElement4.TL4_JP = null;
            isLanguageElement4.TL5_XAA = null;
            this.BTN_CANCEL.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement4});
            this.BTN_CANCEL.Location = new System.Drawing.Point(547, 69);
            this.BTN_CANCEL.Name = "BTN_CANCEL";
            this.BTN_CANCEL.Size = new System.Drawing.Size(137, 21);
            this.BTN_CANCEL.TabIndex = 192;
            this.BTN_CANCEL.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_CANCEL_ButtonClick);
            // 
            // BTN_PDF_IMPORT
            // 
            this.BTN_PDF_IMPORT.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_PDF_IMPORT.ButtonText = "Pdf Import";
            isLanguageElement5.Default = "Pdf Import";
            isLanguageElement5.SiteName = null;
            isLanguageElement5.TL1_KR = "Pdf 등록";
            isLanguageElement5.TL2_CN = null;
            isLanguageElement5.TL3_VN = null;
            isLanguageElement5.TL4_JP = null;
            isLanguageElement5.TL5_XAA = null;
            this.BTN_PDF_IMPORT.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement5});
            this.BTN_PDF_IMPORT.Location = new System.Drawing.Point(547, 42);
            this.BTN_PDF_IMPORT.Name = "BTN_PDF_IMPORT";
            this.BTN_PDF_IMPORT.Size = new System.Drawing.Size(137, 21);
            this.BTN_PDF_IMPORT.TabIndex = 191;
            this.BTN_PDF_IMPORT.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_PDF_IMPORT_ButtonClick);
            // 
            // BTN_FILE_FIND
            // 
            this.BTN_FILE_FIND.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.BTN_FILE_FIND.ButtonText = "File Find...";
            isLanguageElement6.Default = "File Find...";
            isLanguageElement6.SiteName = null;
            isLanguageElement6.TL1_KR = "파일 찾기...";
            isLanguageElement6.TL2_CN = null;
            isLanguageElement6.TL3_VN = null;
            isLanguageElement6.TL4_JP = null;
            isLanguageElement6.TL5_XAA = null;
            this.BTN_FILE_FIND.ButtonTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement6});
            this.BTN_FILE_FIND.Location = new System.Drawing.Point(547, 14);
            this.BTN_FILE_FIND.Name = "BTN_FILE_FIND";
            this.BTN_FILE_FIND.Size = new System.Drawing.Size(137, 21);
            this.BTN_FILE_FIND.TabIndex = 189;
            this.BTN_FILE_FIND.ButtonClick += new InfoSummit.Win.ControlAdv.ISButton.ClickEventHandler(this.BTN_FILE_FIND_ButtonClick);
            // 
            // V_FILE_PWD
            // 
            this.V_FILE_PWD.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.V_FILE_PWD.ComboBoxValue = "";
            this.V_FILE_PWD.ComboData = null;
            this.V_FILE_PWD.CurrencyValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_FILE_PWD.DataAdapter = null;
            this.V_FILE_PWD.DataColumn = null;
            this.V_FILE_PWD.DateTimeValue = new System.DateTime(2010, 3, 17, 0, 0, 0, 0);
            this.V_FILE_PWD.DoubleValue = 0;
            this.V_FILE_PWD.EditValue = "";
            this.V_FILE_PWD.Location = new System.Drawing.Point(7, 69);
            this.V_FILE_PWD.LookupAdapter = null;
            this.V_FILE_PWD.Name = "V_FILE_PWD";
            this.V_FILE_PWD.NumberValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_FILE_PWD.PercentValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.V_FILE_PWD.PromptText = "Pdf Password";
            isLanguageElement7.Default = "Pdf Password";
            isLanguageElement7.SiteName = null;
            isLanguageElement7.TL1_KR = "Pdf 비밀번호";
            isLanguageElement7.TL2_CN = null;
            isLanguageElement7.TL3_VN = null;
            isLanguageElement7.TL4_JP = null;
            isLanguageElement7.TL5_XAA = null;
            this.V_FILE_PWD.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement7});
            this.V_FILE_PWD.Size = new System.Drawing.Size(206, 21);
            this.V_FILE_PWD.TabIndex = 190;
            this.V_FILE_PWD.TextValue = "";
            this.grpUtf8.ResumeLayout(false);
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox1)).EndInit();
            this.isGroupBox1.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpUtf8;
        private System.Windows.Forms.RichTextBox txtUtf8;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_FILE_NAME;
        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISGroupBox isGroupBox1;
        private InfoSummit.Win.ControlAdv.ISButton BTN_FILE_FIND;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_FILE_PWD;
        private InfoSummit.Win.ControlAdv.ISButton BTN_CANCEL;
        private InfoSummit.Win.ControlAdv.ISButton BTN_PDF_IMPORT;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_PERSON_NUM;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_NAME;
        private InfoSummit.Win.ControlAdv.ISEditAdv V_YYYYMM;
    }
}

