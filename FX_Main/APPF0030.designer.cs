namespace FX_Main
{
    partial class APPF0030
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
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement1 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement2 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement3 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement4 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement5 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISLanguageElement isLanguageElement1 = new InfoSummit.Win.ControlAdv.ISLanguageElement();
            Syncfusion.Windows.Forms.Tools.TreeNodeAdvStyleInfo treeNodeAdvStyleInfo1 = new Syncfusion.Windows.Forms.Tools.TreeNodeAdvStyleInfo();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement1 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement2 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement3 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement4 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement5 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement6 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement6 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement7 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement8 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement9 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement10 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement11 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement12 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement13 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraColElement isOraColElement14 = new InfoSummit.Win.ControlAdv.ISOraColElement();
            InfoSummit.Win.ControlAdv.ISOraParamElement isOraParamElement7 = new InfoSummit.Win.ControlAdv.ISOraParamElement();
            this.isAppInterfaceAdv1 = new InfoSummit.Win.ControlAdv.ISAppInterfaceAdv(this.components);
            this.isOraConnection1 = new InfoSummit.Win.ControlAdv.ISOraConnection(this.components);
            this.isDataCommand1 = new InfoSummit.Win.ControlAdv.ISDataCommand(this.components);
            this.isGroupBox1 = new InfoSummit.Win.ControlAdv.ISGroupBox();
            this.gradientLabel1 = new Syncfusion.Windows.Forms.Tools.GradientLabel();
            this.buttonAdv1 = new Syncfusion.Windows.Forms.ButtonAdv();
            this.isTreeView1 = new InfoSummit.Win.ControlAdv.ISTreeView();
            this.idaNavigatorMenuAll = new InfoSummit.Win.ControlAdv.ISDataAdapter(this.components);
            this.idaNavigatorMenuEntryAll = new InfoSummit.Win.ControlAdv.ISDataAdapter(this.components);
            this.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox1)).BeginInit();
            this.isGroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.isTreeView1)).BeginInit();
            // 
            // isOraConnection1
            // 
            this.isOraConnection1.AppInterfaceAdv = this.isAppInterfaceAdv1;
            this.isOraConnection1.OraConnectionInfo = oraConnectionInfo1;
            this.isOraConnection1.OraHost = "112.216.63.194";
            this.isOraConnection1.OraPassword = "infoflex";
            this.isOraConnection1.OraPort = "1521";
            this.isOraConnection1.OraServiceName = "GTCDEV";
            this.isOraConnection1.OraUserId = "apps";
            // 
            // isDataCommand1
            // 
            isOraParamElement1.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement1.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement1.MemberValue = "SOB_ID";
            isOraParamElement1.OraDbTypeString = "NUMBER";
            isOraParamElement1.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement1.ParamName = "W_SOB_ID";
            isOraParamElement1.Size = 22;
            isOraParamElement1.SourceColumn = null;
            isOraParamElement2.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement2.MemberControl = this.isAppInterfaceAdv1;
            isOraParamElement2.MemberValue = "ORG_ID";
            isOraParamElement2.OraDbTypeString = "NUMBER";
            isOraParamElement2.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement2.ParamName = "W_ORG_ID";
            isOraParamElement2.Size = 22;
            isOraParamElement2.SourceColumn = null;
            isOraParamElement3.Direction = System.Data.ParameterDirection.Input;
            isOraParamElement3.MemberControl = null;
            isOraParamElement3.MemberValue = null;
            isOraParamElement3.OraDbTypeString = "NUMBER";
            isOraParamElement3.OraType = System.Data.OracleClient.OracleType.Number;
            isOraParamElement3.ParamName = "W_ASSEMBLY_INFO_ID";
            isOraParamElement3.Size = 22;
            isOraParamElement3.SourceColumn = null;
            isOraParamElement4.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement4.MemberControl = null;
            isOraParamElement4.MemberValue = null;
            isOraParamElement4.OraDbTypeString = "VARCHAR2";
            isOraParamElement4.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement4.ParamName = "O_ASSEMBLY_FILE_NAME";
            isOraParamElement4.Size = 150;
            isOraParamElement4.SourceColumn = null;
            isOraParamElement5.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement5.MemberControl = null;
            isOraParamElement5.MemberValue = null;
            isOraParamElement5.OraDbTypeString = "VARCHAR2";
            isOraParamElement5.OraType = System.Data.OracleClient.OracleType.VarChar;
            isOraParamElement5.ParamName = "O_ASSEMBLY_PATH";
            isOraParamElement5.Size = 254;
            isOraParamElement5.SourceColumn = null;
            this.isDataCommand1.CommandParamElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraParamElement[] {
            isOraParamElement1,
            isOraParamElement2,
            isOraParamElement3,
            isOraParamElement4,
            isOraParamElement5});
            this.isDataCommand1.DataTransaction = null;
            // 
            // APPF0030
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.ClientSize = new System.Drawing.Size(388, 524);
            this.Controls.Add(this.isGroupBox1);
            this.Name = "APPF0030";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Navigator";
            this.Shown += new System.EventHandler(this.APPF0030_Shown);
            this.Load += new System.EventHandler(this.APPF0030_Load);
            this.isDataCommand1.OraConnection = this.isOraConnection1;
            this.isDataCommand1.OraOwner = "APPS";
            this.isDataCommand1.OraPackage = "ZZZZ_DEVELOPER_ASSEMBLY_G";
            this.isDataCommand1.OraProcedure = "ASSEMBLY_PROCESS_START";
            // 
            // isGroupBox1
            // 
            this.isGroupBox1.AppInterfaceAdv = null;
            this.isGroupBox1.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(176)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.isGroupBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.isGroupBox1.Controls.Add(this.gradientLabel1);
            this.isGroupBox1.Controls.Add(this.buttonAdv1);
            this.isGroupBox1.Controls.Add(this.isTreeView1);
            this.isGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.isGroupBox1.Location = new System.Drawing.Point(5, 5);
            this.isGroupBox1.Name = "isGroupBox1";
            this.isGroupBox1.PromptText = "isGroupBox1";
            isLanguageElement1.Default = "isGroupBox1";
            isLanguageElement1.SiteName = null;
            isLanguageElement1.TL1_KR = null;
            isLanguageElement1.TL2_CN = null;
            isLanguageElement1.TL3_VN = null;
            isLanguageElement1.TL4_JP = null;
            isLanguageElement1.TL5_XAA = null;
            this.isGroupBox1.PromptTextElement.AddRange(new InfoSummit.Win.ControlAdv.ISLanguageElement[] {
            isLanguageElement1});
            this.isGroupBox1.PromptVisible = false;
            this.isGroupBox1.Size = new System.Drawing.Size(378, 514);
            this.isGroupBox1.TabIndex = 0;
            // 
            // gradientLabel1
            // 
            this.gradientLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gradientLabel1.BackgroundColor = new Syncfusion.Drawing.BrushInfo(Syncfusion.Drawing.GradientStyle.Vertical, System.Drawing.Color.FromArgb(((int)(((byte)(241)))), ((int)(((byte)(244)))), ((int)(((byte)(254))))), System.Drawing.Color.FromArgb(((int)(((byte)(180)))), ((int)(((byte)(220)))), ((int)(((byte)(250))))));
            this.gradientLabel1.BorderSides = ((System.Windows.Forms.Border3DSide)((((System.Windows.Forms.Border3DSide.Left | System.Windows.Forms.Border3DSide.Top)
                        | System.Windows.Forms.Border3DSide.Right)
                        | System.Windows.Forms.Border3DSide.Bottom)));
            this.gradientLabel1.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust;
            this.gradientLabel1.Font = new System.Drawing.Font("Courier New", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gradientLabel1.Location = new System.Drawing.Point(2, 3);
            this.gradientLabel1.Name = "gradientLabel1";
            this.gradientLabel1.Size = new System.Drawing.Size(337, 32);
            this.gradientLabel1.TabIndex = 193;
            this.gradientLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // buttonAdv1
            // 
            this.buttonAdv1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonAdv1.Appearance = Syncfusion.Windows.Forms.ButtonAppearance.Office2007;
            this.buttonAdv1.ButtonType = Syncfusion.Windows.Forms.Tools.ButtonTypes.Undo;
            this.buttonAdv1.Location = new System.Drawing.Point(342, 3);
            this.buttonAdv1.Name = "buttonAdv1";
            this.buttonAdv1.Size = new System.Drawing.Size(32, 32);
            this.buttonAdv1.TabIndex = 192;
            this.buttonAdv1.Text = "buttonAdv1";
            this.buttonAdv1.UseVisualStyle = true;
            this.buttonAdv1.Click += new System.EventHandler(this.buttonAdv1_Click);
            // 
            // isTreeView1
            // 
            this.isTreeView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            treeNodeAdvStyleInfo1.EnsureDefaultOptionedChild = true;
            this.isTreeView1.BaseStylePairs.AddRange(new Syncfusion.Windows.Forms.Tools.StyleNamePair[] {
            new Syncfusion.Windows.Forms.Tools.StyleNamePair("Standard", treeNodeAdvStyleInfo1)});
            this.isTreeView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.isTreeView1.Font = new System.Drawing.Font("돋움", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            // 
            // 
            // 
            this.isTreeView1.HelpTextControl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.isTreeView1.HelpTextControl.Location = new System.Drawing.Point(0, 0);
            this.isTreeView1.HelpTextControl.Name = "helpText";
            this.isTreeView1.HelpTextControl.Size = new System.Drawing.Size(55, 14);
            this.isTreeView1.HelpTextControl.TabIndex = 0;
            this.isTreeView1.HelpTextControl.Text = "help text";
            this.isTreeView1.Location = new System.Drawing.Point(3, 37);
            this.isTreeView1.Name = "isTreeView1";
            this.isTreeView1.Office2007ScrollBars = true;
            this.isTreeView1.Office2007ScrollBarsColorScheme = Syncfusion.Windows.Forms.Office2007ColorScheme.Managed;
            this.isTreeView1.Size = new System.Drawing.Size(372, 474);
            this.isTreeView1.TabIndex = 190;
            this.isTreeView1.Text = "isTreeView1";
            // 
            // 
            // 
            this.isTreeView1.ToolTipControl.BackColor = System.Drawing.SystemColors.Info;
            this.isTreeView1.ToolTipControl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.isTreeView1.ToolTipControl.Location = new System.Drawing.Point(0, 0);
            this.isTreeView1.ToolTipControl.Name = "toolTip";
            this.isTreeView1.ToolTipControl.Size = new System.Drawing.Size(45, 14);
            this.isTreeView1.ToolTipControl.TabIndex = 1;
            this.isTreeView1.ToolTipControl.Text = "toolTip";
            this.isTreeView1.DoubleClick += new System.EventHandler(this.isTreeView1_DoubleClick);
            // 
            // idaNavigatorMenuAll
            // 
            this.idaNavigatorMenuAll.CancelMember.Cancel = false;
            this.idaNavigatorMenuAll.CancelMember.Member = null;
            this.idaNavigatorMenuAll.CancelMember.Prompt = null;
            this.idaNavigatorMenuAll.CancelMember.TabIndex = -1;
            this.idaNavigatorMenuAll.CancelMember.ValueItem = null;
            this.idaNavigatorMenuAll.CancelUpdateFilterString = null;
            this.idaNavigatorMenuAll.CancelUpdateRow = null;
            this.idaNavigatorMenuAll.DataTransaction = null;
            this.idaNavigatorMenuAll.FocusedControl = null;
            this.idaNavigatorMenuAll.MasterAdapter = null;
            this.idaNavigatorMenuAll.OraConnection = this.isOraConnection1;
            this.idaNavigatorMenuAll.OraDelete = null;
            this.idaNavigatorMenuAll.OraInsert = null;
            this.idaNavigatorMenuAll.OraOwner = "APPS";
            this.idaNavigatorMenuAll.OraPackage = "ZZZZ_DEVELOPER_MENU_G";
            this.idaNavigatorMenuAll.OraSelect = "NAVI_MENU_SIMPLY";
            this.idaNavigatorMenuAll.OraSelectData = null;
            this.idaNavigatorMenuAll.OraUpdate = null;
            isOraColElement1.DataColumn = "MENU_ID";
            isOraColElement1.DataOrdinal = 0;
            isOraColElement1.DataType = "System.Decimal";
            isOraColElement1.HeaderPrompt = "MENU_ID";
            isOraColElement1.LastValue = null;
            isOraColElement1.MemberControl = null;
            isOraColElement1.MemberValue = null;
            isOraColElement1.Nullable = 0;
            isOraColElement1.Ordinal = 0;
            isOraColElement1.RelationKeyColumn = null;
            isOraColElement1.ReturnParameter = null;
            isOraColElement1.TL1_KR = null;
            isOraColElement1.TL2_CN = null;
            isOraColElement1.TL3_VN = null;
            isOraColElement1.TL4_JP = null;
            isOraColElement1.TL5_XAA = null;
            isOraColElement1.Visible = null;
            isOraColElement1.Width = null;
            isOraColElement2.DataColumn = "MENU_SEQ";
            isOraColElement2.DataOrdinal = 1;
            isOraColElement2.DataType = "System.Decimal";
            isOraColElement2.HeaderPrompt = "MENU_SEQ";
            isOraColElement2.LastValue = null;
            isOraColElement2.MemberControl = null;
            isOraColElement2.MemberValue = null;
            isOraColElement2.Nullable = 1;
            isOraColElement2.Ordinal = 1;
            isOraColElement2.RelationKeyColumn = null;
            isOraColElement2.ReturnParameter = null;
            isOraColElement2.TL1_KR = null;
            isOraColElement2.TL2_CN = null;
            isOraColElement2.TL3_VN = null;
            isOraColElement2.TL4_JP = null;
            isOraColElement2.TL5_XAA = null;
            isOraColElement2.Visible = null;
            isOraColElement2.Width = null;
            isOraColElement3.DataColumn = "MENU_NAME";
            isOraColElement3.DataOrdinal = 2;
            isOraColElement3.DataType = "System.String";
            isOraColElement3.HeaderPrompt = "MENU_NAME";
            isOraColElement3.LastValue = null;
            isOraColElement3.MemberControl = null;
            isOraColElement3.MemberValue = null;
            isOraColElement3.Nullable = 0;
            isOraColElement3.Ordinal = 2;
            isOraColElement3.RelationKeyColumn = null;
            isOraColElement3.ReturnParameter = null;
            isOraColElement3.TL1_KR = null;
            isOraColElement3.TL2_CN = null;
            isOraColElement3.TL3_VN = null;
            isOraColElement3.TL4_JP = null;
            isOraColElement3.TL5_XAA = null;
            isOraColElement3.Visible = null;
            isOraColElement3.Width = null;
            isOraColElement4.DataColumn = "MENU_PROMPT";
            isOraColElement4.DataOrdinal = 3;
            isOraColElement4.DataType = "System.String";
            isOraColElement4.HeaderPrompt = "MENU_PROMPT";
            isOraColElement4.LastValue = null;
            isOraColElement4.MemberControl = null;
            isOraColElement4.MemberValue = null;
            isOraColElement4.Nullable = 0;
            isOraColElement4.Ordinal = 3;
            isOraColElement4.RelationKeyColumn = null;
            isOraColElement4.ReturnParameter = null;
            isOraColElement4.TL1_KR = null;
            isOraColElement4.TL2_CN = null;
            isOraColElement4.TL3_VN = null;
            isOraColElement4.TL4_JP = null;
            isOraColElement4.TL5_XAA = null;
            isOraColElement4.Visible = null;
            isOraColElement4.Width = null;
            isOraColElement5.DataColumn = "MENU_DESC";
            isOraColElement5.DataOrdinal = 4;
            isOraColElement5.DataType = "System.String";
            isOraColElement5.HeaderPrompt = "MENU_DESC";
            isOraColElement5.LastValue = null;
            isOraColElement5.MemberControl = null;
            isOraColElement5.MemberValue = null;
            isOraColElement5.Nullable = 1;
            isOraColElement5.Ordinal = 4;
            isOraColElement5.RelationKeyColumn = null;
            isOraColElement5.ReturnParameter = null;
            isOraColElement5.TL1_KR = null;
            isOraColElement5.TL2_CN = null;
            isOraColElement5.TL3_VN = null;
            isOraColElement5.TL4_JP = null;
            isOraColElement5.TL5_XAA = null;
            isOraColElement5.Visible = null;
            isOraColElement5.Width = null;
            this.idaNavigatorMenuAll.SelectColElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraColElement[] {
            isOraColElement1,
            isOraColElement2,
            isOraColElement3,
            isOraColElement4,
            isOraColElement5});
            isOraParamElement6.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement6.MemberControl = null;
            isOraParamElement6.MemberValue = null;
            isOraParamElement6.OraDbTypeString = "REF CURSOR";
            isOraParamElement6.OraType = System.Data.OracleClient.OracleType.Cursor;
            isOraParamElement6.ParamName = "P_CURSOR";
            isOraParamElement6.Size = 0;
            isOraParamElement6.SourceColumn = null;
            this.idaNavigatorMenuAll.SelectParamElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraParamElement[] {
            isOraParamElement6});
            this.idaNavigatorMenuAll.WizardOwner = "";
            this.idaNavigatorMenuAll.WizardProcedure = "";
            this.idaNavigatorMenuAll.WizardTableName = "";
            // 
            // idaNavigatorMenuEntryAll
            // 
            this.idaNavigatorMenuEntryAll.CancelMember.Cancel = false;
            this.idaNavigatorMenuEntryAll.CancelMember.Member = null;
            this.idaNavigatorMenuEntryAll.CancelMember.Prompt = null;
            this.idaNavigatorMenuEntryAll.CancelMember.TabIndex = -1;
            this.idaNavigatorMenuEntryAll.CancelMember.ValueItem = null;
            this.idaNavigatorMenuEntryAll.CancelUpdateFilterString = null;
            this.idaNavigatorMenuEntryAll.CancelUpdateRow = null;
            this.idaNavigatorMenuEntryAll.DataTransaction = null;
            this.idaNavigatorMenuEntryAll.FocusedControl = null;
            this.idaNavigatorMenuEntryAll.MasterAdapter = null;
            this.idaNavigatorMenuEntryAll.OraConnection = this.isOraConnection1;
            this.idaNavigatorMenuEntryAll.OraDelete = null;
            this.idaNavigatorMenuEntryAll.OraInsert = null;
            this.idaNavigatorMenuEntryAll.OraOwner = "APPS";
            this.idaNavigatorMenuEntryAll.OraPackage = "ZZZZ_DEVELOPER_MENU_G";
            this.idaNavigatorMenuEntryAll.OraSelect = "NAVI_MENU_ENTRY_SIMPLY";
            this.idaNavigatorMenuEntryAll.OraSelectData = null;
            this.idaNavigatorMenuEntryAll.OraUpdate = null;
            isOraColElement6.DataColumn = "ENTRY_ID";
            isOraColElement6.DataOrdinal = 0;
            isOraColElement6.DataType = "System.Decimal";
            isOraColElement6.HeaderPrompt = "ENTRY_ID";
            isOraColElement6.LastValue = null;
            isOraColElement6.MemberControl = null;
            isOraColElement6.MemberValue = null;
            isOraColElement6.Nullable = 0;
            isOraColElement6.Ordinal = 0;
            isOraColElement6.RelationKeyColumn = null;
            isOraColElement6.ReturnParameter = null;
            isOraColElement6.TL1_KR = null;
            isOraColElement6.TL2_CN = null;
            isOraColElement6.TL3_VN = null;
            isOraColElement6.TL4_JP = null;
            isOraColElement6.TL5_XAA = null;
            isOraColElement6.Visible = null;
            isOraColElement6.Width = null;
            isOraColElement7.DataColumn = "MENU_ID";
            isOraColElement7.DataOrdinal = 1;
            isOraColElement7.DataType = "System.Decimal";
            isOraColElement7.HeaderPrompt = "MENU_ID";
            isOraColElement7.LastValue = null;
            isOraColElement7.MemberControl = null;
            isOraColElement7.MemberValue = null;
            isOraColElement7.Nullable = 0;
            isOraColElement7.Ordinal = 1;
            isOraColElement7.RelationKeyColumn = null;
            isOraColElement7.ReturnParameter = null;
            isOraColElement7.TL1_KR = null;
            isOraColElement7.TL2_CN = null;
            isOraColElement7.TL3_VN = null;
            isOraColElement7.TL4_JP = null;
            isOraColElement7.TL5_XAA = null;
            isOraColElement7.Visible = null;
            isOraColElement7.Width = null;
            isOraColElement8.DataColumn = "ENTRY_SEQ";
            isOraColElement8.DataOrdinal = 2;
            isOraColElement8.DataType = "System.Decimal";
            isOraColElement8.HeaderPrompt = "ENTRY_SEQ";
            isOraColElement8.LastValue = null;
            isOraColElement8.MemberControl = null;
            isOraColElement8.MemberValue = null;
            isOraColElement8.Nullable = 0;
            isOraColElement8.Ordinal = 2;
            isOraColElement8.RelationKeyColumn = null;
            isOraColElement8.ReturnParameter = null;
            isOraColElement8.TL1_KR = null;
            isOraColElement8.TL2_CN = null;
            isOraColElement8.TL3_VN = null;
            isOraColElement8.TL4_JP = null;
            isOraColElement8.TL5_XAA = null;
            isOraColElement8.Visible = null;
            isOraColElement8.Width = null;
            isOraColElement9.DataColumn = "ENTRY_PROMPT";
            isOraColElement9.DataOrdinal = 3;
            isOraColElement9.DataType = "System.String";
            isOraColElement9.HeaderPrompt = "ENTRY_PROMPT";
            isOraColElement9.LastValue = null;
            isOraColElement9.MemberControl = null;
            isOraColElement9.MemberValue = null;
            isOraColElement9.Nullable = 0;
            isOraColElement9.Ordinal = 3;
            isOraColElement9.RelationKeyColumn = null;
            isOraColElement9.ReturnParameter = null;
            isOraColElement9.TL1_KR = null;
            isOraColElement9.TL2_CN = null;
            isOraColElement9.TL3_VN = null;
            isOraColElement9.TL4_JP = null;
            isOraColElement9.TL5_XAA = null;
            isOraColElement9.Visible = null;
            isOraColElement9.Width = null;
            isOraColElement10.DataColumn = "SUB_MENU_ID";
            isOraColElement10.DataOrdinal = 4;
            isOraColElement10.DataType = "System.Decimal";
            isOraColElement10.HeaderPrompt = "SUB_MENU_ID";
            isOraColElement10.LastValue = null;
            isOraColElement10.MemberControl = null;
            isOraColElement10.MemberValue = null;
            isOraColElement10.Nullable = 1;
            isOraColElement10.Ordinal = 4;
            isOraColElement10.RelationKeyColumn = null;
            isOraColElement10.ReturnParameter = null;
            isOraColElement10.TL1_KR = null;
            isOraColElement10.TL2_CN = null;
            isOraColElement10.TL3_VN = null;
            isOraColElement10.TL4_JP = null;
            isOraColElement10.TL5_XAA = null;
            isOraColElement10.Visible = null;
            isOraColElement10.Width = null;
            isOraColElement11.DataColumn = "ASSEMBLY_INFO_ID";
            isOraColElement11.DataOrdinal = 5;
            isOraColElement11.DataType = "System.Decimal";
            isOraColElement11.HeaderPrompt = "ASSEMBLY_INFO_ID";
            isOraColElement11.LastValue = null;
            isOraColElement11.MemberControl = null;
            isOraColElement11.MemberValue = null;
            isOraColElement11.Nullable = 1;
            isOraColElement11.Ordinal = 5;
            isOraColElement11.RelationKeyColumn = null;
            isOraColElement11.ReturnParameter = null;
            isOraColElement11.TL1_KR = null;
            isOraColElement11.TL2_CN = null;
            isOraColElement11.TL3_VN = null;
            isOraColElement11.TL4_JP = null;
            isOraColElement11.TL5_XAA = null;
            isOraColElement11.Visible = null;
            isOraColElement11.Width = null;
            isOraColElement12.DataColumn = "ENTRY_DESC";
            isOraColElement12.DataOrdinal = 6;
            isOraColElement12.DataType = "System.String";
            isOraColElement12.HeaderPrompt = "ENTRY_DESC";
            isOraColElement12.LastValue = null;
            isOraColElement12.MemberControl = null;
            isOraColElement12.MemberValue = null;
            isOraColElement12.Nullable = 1;
            isOraColElement12.Ordinal = 6;
            isOraColElement12.RelationKeyColumn = null;
            isOraColElement12.ReturnParameter = null;
            isOraColElement12.TL1_KR = null;
            isOraColElement12.TL2_CN = null;
            isOraColElement12.TL3_VN = null;
            isOraColElement12.TL4_JP = null;
            isOraColElement12.TL5_XAA = null;
            isOraColElement12.Visible = null;
            isOraColElement12.Width = null;
            isOraColElement13.DataColumn = "ASSEMBLY_ID";
            isOraColElement13.DataOrdinal = 7;
            isOraColElement13.DataType = "System.String";
            isOraColElement13.HeaderPrompt = "ASSEMBLY_ID";
            isOraColElement13.LastValue = null;
            isOraColElement13.MemberControl = null;
            isOraColElement13.MemberValue = null;
            isOraColElement13.Nullable = 1;
            isOraColElement13.Ordinal = 7;
            isOraColElement13.RelationKeyColumn = null;
            isOraColElement13.ReturnParameter = null;
            isOraColElement13.TL1_KR = null;
            isOraColElement13.TL2_CN = null;
            isOraColElement13.TL3_VN = null;
            isOraColElement13.TL4_JP = null;
            isOraColElement13.TL5_XAA = null;
            isOraColElement13.Visible = null;
            isOraColElement13.Width = null;
            isOraColElement14.DataColumn = "ASSEMBLY_FILE_NAME";
            isOraColElement14.DataOrdinal = 8;
            isOraColElement14.DataType = "System.String";
            isOraColElement14.HeaderPrompt = "ASSEMBLY_FILE_NAME";
            isOraColElement14.LastValue = null;
            isOraColElement14.MemberControl = null;
            isOraColElement14.MemberValue = null;
            isOraColElement14.Nullable = 1;
            isOraColElement14.Ordinal = 8;
            isOraColElement14.RelationKeyColumn = null;
            isOraColElement14.ReturnParameter = null;
            isOraColElement14.TL1_KR = null;
            isOraColElement14.TL2_CN = null;
            isOraColElement14.TL3_VN = null;
            isOraColElement14.TL4_JP = null;
            isOraColElement14.TL5_XAA = null;
            isOraColElement14.Visible = null;
            isOraColElement14.Width = null;
            this.idaNavigatorMenuEntryAll.SelectColElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraColElement[] {
            isOraColElement6,
            isOraColElement7,
            isOraColElement8,
            isOraColElement9,
            isOraColElement10,
            isOraColElement11,
            isOraColElement12,
            isOraColElement13,
            isOraColElement14});
            isOraParamElement7.Direction = System.Data.ParameterDirection.Output;
            isOraParamElement7.MemberControl = null;
            isOraParamElement7.MemberValue = null;
            isOraParamElement7.OraDbTypeString = "REF CURSOR";
            isOraParamElement7.OraType = System.Data.OracleClient.OracleType.Cursor;
            isOraParamElement7.ParamName = "P_CURSOR";
            isOraParamElement7.Size = 0;
            isOraParamElement7.SourceColumn = null;
            this.idaNavigatorMenuEntryAll.SelectParamElement.AddRange(new InfoSummit.Win.ControlAdv.ISOraParamElement[] {
            isOraParamElement7});
            this.idaNavigatorMenuEntryAll.WizardOwner = "";
            this.idaNavigatorMenuEntryAll.WizardProcedure = "";
            this.idaNavigatorMenuEntryAll.WizardTableName = "";
            this.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.isGroupBox1)).EndInit();
            this.isGroupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.isTreeView1)).EndInit();

        }

        #endregion

        private InfoSummit.Win.ControlAdv.ISOraConnection isOraConnection1;
        private InfoSummit.Win.ControlAdv.ISAppInterfaceAdv isAppInterfaceAdv1;
        private InfoSummit.Win.ControlAdv.ISDataCommand isDataCommand1;
        private InfoSummit.Win.ControlAdv.ISGroupBox isGroupBox1;
        private Syncfusion.Windows.Forms.Tools.GradientLabel gradientLabel1;
        private Syncfusion.Windows.Forms.ButtonAdv buttonAdv1;
        private InfoSummit.Win.ControlAdv.ISTreeView isTreeView1;
        private InfoSummit.Win.ControlAdv.ISDataAdapter idaNavigatorMenuAll;
        private InfoSummit.Win.ControlAdv.ISDataAdapter idaNavigatorMenuEntryAll;
    }
}