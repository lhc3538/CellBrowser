namespace MyBrowser
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false </param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.MenuStrip = new System.Windows.Forms.MenuStrip();
            this.MenuFile = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuNew = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuNewTab = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuOpen = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuSave = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.MenuPrint = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuPrintPreview = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuSetting = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.MenuProperty = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuQuit = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuView = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuToolBar = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuAddressBar = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuGoto = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuBack = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuForward = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuStop = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuRefresh = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuSource = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuTool = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuSnapshot = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolBar = new System.Windows.Forms.ToolStrip();
            this.ToolBack = new System.Windows.Forms.ToolStripSplitButton();
            this.ToolForward = new System.Windows.Forms.ToolStripSplitButton();
            this.ToolStop = new System.Windows.Forms.ToolStripButton();
            this.ToolRefresh = new System.Windows.Forms.ToolStripButton();
            this.ToolHome = new System.Windows.Forms.ToolStripButton();
            this.ToolFavorites = new System.Windows.Forms.ToolStripSplitButton();
            this.MenuFavorites = new System.Windows.Forms.ToolStripMenuItem();
            this.AddressBar = new System.Windows.Forms.ToolStrip();
            this.ToolLabel = new System.Windows.Forms.ToolStripLabel();
            this.ToolAddress = new System.Windows.Forms.ToolStripComboBox();
            this.ToolGoto = new System.Windows.Forms.ToolStripButton();
            this.StatusStrip = new System.Windows.Forms.StatusStrip();
            this.ToolImage = new System.Windows.Forms.ToolStripStatusLabel();
            this.StatusText = new System.Windows.Forms.ToolStripStatusLabel();
            this.ToolProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.ToolInternet = new System.Windows.Forms.ToolStripStatusLabel();
            this.TabControl = new System.Windows.Forms.TabControl();
            this.MenuStrip.SuspendLayout();
            this.ToolBar.SuspendLayout();
            this.AddressBar.SuspendLayout();
            this.StatusStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // MenuStrip
            // 
            this.MenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuFile,
            this.MenuView,
            this.MenuTool,
            this.MenuHelp});
            this.MenuStrip.Location = new System.Drawing.Point(0, 0);
            this.MenuStrip.Name = "MenuStrip";
            this.MenuStrip.Size = new System.Drawing.Size(781, 24);
            this.MenuStrip.TabIndex = 0;
            // 
            // MenuFile
            // 
            this.MenuFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuNew,
            this.MenuNewTab,
            this.MenuOpen,
            this.MenuSave,
            this.MenuSeparator1,
            this.MenuPrint,
            this.MenuPrintPreview,
            this.MenuSetting,
            this.MenuSeparator2,
            this.MenuProperty,
            this.MenuQuit});
            this.MenuFile.Name = "MenuFile";
            this.MenuFile.Size = new System.Drawing.Size(59, 20);
            this.MenuFile.Text = "文件(&F)";
            // 
            // MenuNew
            // 
            this.MenuNew.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.MenuNew.Name = "MenuNew";
            this.MenuNew.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.MenuNew.Size = new System.Drawing.Size(177, 22);
            this.MenuNew.Text = "新建(&N)";
            this.MenuNew.Click += new System.EventHandler(this.MenuNew_Click);
            // 
            // MenuNewTab
            // 
            this.MenuNewTab.Name = "MenuNewTab";
            this.MenuNewTab.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.T)));
            this.MenuNewTab.Size = new System.Drawing.Size(177, 22);
            this.MenuNewTab.Text = "新建标签(&T)";
            this.MenuNewTab.Click += new System.EventHandler(this.MenuNewTab_Click);
            // 
            // MenuOpen
            // 
            this.MenuOpen.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.MenuOpen.Name = "MenuOpen";
            this.MenuOpen.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.MenuOpen.Size = new System.Drawing.Size(177, 22);
            this.MenuOpen.Text = "打开(&O)";
            this.MenuOpen.Click += new System.EventHandler(this.MenuOpen_Click);
            // 
            // MenuSave
            // 
            this.MenuSave.Name = "MenuSave";
            this.MenuSave.ShortcutKeyDisplayString = global::MyBrowser.Properties.Resources.String;
            this.MenuSave.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.MenuSave.Size = new System.Drawing.Size(177, 22);
            this.MenuSave.Text = "另存为(&S)";
            this.MenuSave.Click += new System.EventHandler(this.MenuSave_Click);
            // 
            // MenuSeparator1
            // 
            this.MenuSeparator1.Name = "MenuSeparator1";
            this.MenuSeparator1.Size = new System.Drawing.Size(174, 6);
            // 
            // MenuPrint
            // 
            this.MenuPrint.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.MenuPrint.Name = "MenuPrint";
            this.MenuPrint.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P)));
            this.MenuPrint.Size = new System.Drawing.Size(177, 22);
            this.MenuPrint.Text = "打印(&P)";
            this.MenuPrint.Click += new System.EventHandler(this.MenuPrint_Click);
            // 
            // MenuPrintPreview
            // 
            this.MenuPrintPreview.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.MenuPrintPreview.Name = "MenuPrintPreview";
            this.MenuPrintPreview.Size = new System.Drawing.Size(177, 22);
            this.MenuPrintPreview.Text = "打印预览(&V)";
            this.MenuPrintPreview.Click += new System.EventHandler(this.MenuPrintPreview_Click);
            // 
            // MenuSetting
            // 
            this.MenuSetting.Name = "MenuSetting";
            this.MenuSetting.Size = new System.Drawing.Size(177, 22);
            this.MenuSetting.Text = "页面设置(&U)";
            this.MenuSetting.Click += new System.EventHandler(this.MenuSetting_Click);
            // 
            // MenuSeparator2
            // 
            this.MenuSeparator2.Name = "MenuSeparator2";
            this.MenuSeparator2.Size = new System.Drawing.Size(174, 6);
            // 
            // MenuProperty
            // 
            this.MenuProperty.Name = "MenuProperty";
            this.MenuProperty.Size = new System.Drawing.Size(177, 22);
            this.MenuProperty.Text = "属性(&R)";
            this.MenuProperty.Click += new System.EventHandler(this.MenuProperty_Click);
            // 
            // MenuQuit
            // 
            this.MenuQuit.Name = "MenuQuit";
            this.MenuQuit.Size = new System.Drawing.Size(177, 22);
            this.MenuQuit.Text = "退出(&X)";
            this.MenuQuit.Click += new System.EventHandler(this.MenuQuit_Click);
            // 
            // MenuView
            // 
            this.MenuView.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuToolBar,
            this.MenuAddressBar,
            this.MenuGoto,
            this.MenuStop,
            this.MenuRefresh,
            this.MenuSource});
            this.MenuView.Name = "MenuView";
            this.MenuView.Size = new System.Drawing.Size(59, 20);
            this.MenuView.Text = "查看(&V)";
            // 
            // MenuToolBar
            // 
            this.MenuToolBar.Checked = true;
            this.MenuToolBar.CheckState = System.Windows.Forms.CheckState.Checked;
            this.MenuToolBar.Name = "MenuToolBar";
            this.MenuToolBar.Size = new System.Drawing.Size(129, 22);
            this.MenuToolBar.Text = "工具栏";
            this.MenuToolBar.Click += new System.EventHandler(this.MenuToolBar_Click);
            // 
            // MenuAddressBar
            // 
            this.MenuAddressBar.Checked = true;
            this.MenuAddressBar.CheckState = System.Windows.Forms.CheckState.Checked;
            this.MenuAddressBar.Name = "MenuAddressBar";
            this.MenuAddressBar.Size = new System.Drawing.Size(129, 22);
            this.MenuAddressBar.Text = "地址栏";
            this.MenuAddressBar.Click += new System.EventHandler(this.MenuAddressBar_Click);
            // 
            // MenuGoto
            // 
            this.MenuGoto.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuBack,
            this.MenuForward});
            this.MenuGoto.Name = "MenuGoto";
            this.MenuGoto.Size = new System.Drawing.Size(129, 22);
            this.MenuGoto.Text = "转到(&G)";
            // 
            // MenuBack
            // 
            this.MenuBack.Image = ((System.Drawing.Image)(resources.GetObject("MenuBack.Image")));
            this.MenuBack.Name = "MenuBack";
            this.MenuBack.Size = new System.Drawing.Size(94, 22);
            this.MenuBack.Text = "后退";
            this.MenuBack.Click += new System.EventHandler(this.MenuBack_Click);
            // 
            // MenuForward
            // 
            this.MenuForward.Image = ((System.Drawing.Image)(resources.GetObject("MenuForward.Image")));
            this.MenuForward.Name = "MenuForward";
            this.MenuForward.Size = new System.Drawing.Size(94, 22);
            this.MenuForward.Text = "前进";
            this.MenuForward.Click += new System.EventHandler(this.MenuForward_Click);
            // 
            // MenuStop
            // 
            this.MenuStop.Name = "MenuStop";
            this.MenuStop.Size = new System.Drawing.Size(129, 22);
            this.MenuStop.Text = "停止(&P)";
            this.MenuStop.Click += new System.EventHandler(this.MenuStop_Click);
            // 
            // MenuRefresh
            // 
            this.MenuRefresh.Name = "MenuRefresh";
            this.MenuRefresh.ShortcutKeys = System.Windows.Forms.Keys.F5;
            this.MenuRefresh.Size = new System.Drawing.Size(129, 22);
            this.MenuRefresh.Text = "刷新(&R)";
            this.MenuRefresh.Click += new System.EventHandler(this.MenuRefresh_Click);
            // 
            // MenuSource
            // 
            this.MenuSource.Name = "MenuSource";
            this.MenuSource.Size = new System.Drawing.Size(129, 22);
            this.MenuSource.Text = "源文件(&C)";
            this.MenuSource.Click += new System.EventHandler(this.MenuSource_Click);
            // 
            // MenuTool
            // 
            this.MenuTool.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuSnapshot});
            this.MenuTool.Name = "MenuTool";
            this.MenuTool.Size = new System.Drawing.Size(59, 20);
            this.MenuTool.Text = "工具(&T)";
            // 
            // MenuSnapshot
            // 
            this.MenuSnapshot.Name = "MenuSnapshot";
            this.MenuSnapshot.Size = new System.Drawing.Size(152, 22);
            this.MenuSnapshot.Text = "网页截图(&S)";
            this.MenuSnapshot.Click += new System.EventHandler(this.MenuSnapshot_Click);
            // 
            // MenuHelp
            // 
            this.MenuHelp.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuAbout});
            this.MenuHelp.Name = "MenuHelp";
            this.MenuHelp.Size = new System.Drawing.Size(59, 20);
            this.MenuHelp.Text = "帮助(&H)";
            // 
            // MenuAbout
            // 
            this.MenuAbout.Name = "MenuAbout";
            this.MenuAbout.Size = new System.Drawing.Size(112, 22);
            this.MenuAbout.Text = "关于(&A)";
            this.MenuAbout.Click += new System.EventHandler(this.MenuAbout_Click);
            // 
            // ToolBar
            // 
            this.ToolBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolBack,
            this.ToolForward,
            this.ToolStop,
            this.ToolRefresh,
            this.ToolHome,
            this.ToolFavorites});
            this.ToolBar.Location = new System.Drawing.Point(0, 24);
            this.ToolBar.Name = "ToolBar";
            this.ToolBar.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.ToolBar.Size = new System.Drawing.Size(781, 33);
            this.ToolBar.TabIndex = 1;
            // 
            // ToolBack
            // 
            this.ToolBack.Image = global::MyBrowser.Properties.Resources.Back;
            this.ToolBack.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ToolBack.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ToolBack.Name = "ToolBack";
            this.ToolBack.Size = new System.Drawing.Size(71, 30);
            this.ToolBack.Text = "后退";
            this.ToolBack.Click += new System.EventHandler(this.Link_Click);
            // 
            // ToolForward
            // 
            this.ToolForward.Image = global::MyBrowser.Properties.Resources.Forward;
            this.ToolForward.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ToolForward.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ToolForward.Name = "ToolForward";
            this.ToolForward.Size = new System.Drawing.Size(71, 30);
            this.ToolForward.Text = "前进";
            this.ToolForward.ToolTipText = "前进";
            this.ToolForward.Click += new System.EventHandler(this.Link_Click);
            // 
            // ToolStop
            // 
            this.ToolStop.Image = global::MyBrowser.Properties.Resources.Stop;
            this.ToolStop.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ToolStop.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ToolStop.Margin = new System.Windows.Forms.Padding(0, 1, 0, 0);
            this.ToolStop.Name = "ToolStop";
            this.ToolStop.Size = new System.Drawing.Size(59, 32);
            this.ToolStop.Text = "停止";
            this.ToolStop.Click += new System.EventHandler(this.Link_Click);
            // 
            // ToolRefresh
            // 
            this.ToolRefresh.Image = global::MyBrowser.Properties.Resources.Refresh;
            this.ToolRefresh.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ToolRefresh.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ToolRefresh.Margin = new System.Windows.Forms.Padding(0, 1, 0, 0);
            this.ToolRefresh.Name = "ToolRefresh";
            this.ToolRefresh.Size = new System.Drawing.Size(59, 32);
            this.ToolRefresh.Text = "刷新";
            this.ToolRefresh.Click += new System.EventHandler(this.Link_Click);
            // 
            // ToolHome
            // 
            this.ToolHome.Image = global::MyBrowser.Properties.Resources.Home;
            this.ToolHome.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ToolHome.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ToolHome.Margin = new System.Windows.Forms.Padding(0, 1, 0, 0);
            this.ToolHome.Name = "ToolHome";
            this.ToolHome.Size = new System.Drawing.Size(59, 32);
            this.ToolHome.Text = "主页";
            this.ToolHome.Click += new System.EventHandler(this.ToolHome_Click);
            // 
            // ToolFavorites
            // 
            this.ToolFavorites.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuFavorites});
            this.ToolFavorites.Image = global::MyBrowser.Properties.Resources.Favorites;
            this.ToolFavorites.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ToolFavorites.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ToolFavorites.Name = "ToolFavorites";
            this.ToolFavorites.Size = new System.Drawing.Size(83, 30);
            this.ToolFavorites.Text = "收藏夹";
            // 
            // MenuFavorites
            // 
            this.MenuFavorites.Name = "MenuFavorites";
            this.MenuFavorites.Size = new System.Drawing.Size(178, 22);
            this.MenuFavorites.Text = "添加到收藏夹(&A)...";
            this.MenuFavorites.Click += new System.EventHandler(this.MenuFavorites_Click);
            // 
            // AddressBar
            // 
            this.AddressBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolLabel,
            this.ToolAddress,
            this.ToolGoto});
            this.AddressBar.Location = new System.Drawing.Point(0, 57);
            this.AddressBar.Name = "AddressBar";
            this.AddressBar.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.AddressBar.Size = new System.Drawing.Size(781, 27);
            this.AddressBar.Stretch = true;
            this.AddressBar.TabIndex = 2;
            this.AddressBar.SizeChanged += new System.EventHandler(this.AddressBar_SizeChanged);
            // 
            // ToolLabel
            // 
            this.ToolLabel.Name = "ToolLabel";
            this.ToolLabel.Size = new System.Drawing.Size(29, 24);
            this.ToolLabel.Text = "地址";
            // 
            // ToolAddress
            // 
            this.ToolAddress.Name = "ToolAddress";
            this.ToolAddress.Size = new System.Drawing.Size(650, 27);
            this.ToolAddress.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ToolAddress_KeyDown);
            // 
            // ToolGoto
            // 
            this.ToolGoto.Image = global::MyBrowser.Properties.Resources.Go;
            this.ToolGoto.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ToolGoto.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ToolGoto.Name = "ToolGoto";
            this.ToolGoto.Size = new System.Drawing.Size(53, 24);
            this.ToolGoto.Text = "转到";
            this.ToolGoto.Click += new System.EventHandler(this.ToolGoto_Click);
            // 
            // StatusStrip
            // 
            this.StatusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolImage,
            this.StatusText,
            this.ToolProgressBar,
            this.ToolInternet});
            this.StatusStrip.Location = new System.Drawing.Point(0, 429);
            this.StatusStrip.Name = "StatusStrip";
            this.StatusStrip.Size = new System.Drawing.Size(781, 22);
            this.StatusStrip.TabIndex = 4;
            // 
            // ToolImage
            // 
            this.ToolImage.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ToolImage.Image = global::MyBrowser.Properties.Resources.Link;
            this.ToolImage.Name = "ToolImage";
            this.ToolImage.Size = new System.Drawing.Size(16, 17);
            // 
            // StatusText
            // 
            this.StatusText.AutoSize = false;
            this.StatusText.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
            this.StatusText.Name = "StatusText";
            this.StatusText.Size = new System.Drawing.Size(632, 17);
            this.StatusText.Spring = true;
            this.StatusText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // ToolProgressBar
            // 
            this.ToolProgressBar.Name = "ToolProgressBar";
            this.ToolProgressBar.Size = new System.Drawing.Size(100, 16);
            // 
            // ToolInternet
            // 
            this.ToolInternet.Image = global::MyBrowser.Properties.Resources.Internet;
            this.ToolInternet.Name = "ToolInternet";
            this.ToolInternet.Size = new System.Drawing.Size(16, 17);
            // 
            // TabControl
            // 
            this.TabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TabControl.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TabControl.ItemSize = new System.Drawing.Size(120, 25);
            this.TabControl.Location = new System.Drawing.Point(0, 84);
            this.TabControl.Name = "TabControl";
            this.TabControl.SelectedIndex = 0;
            this.TabControl.ShowToolTips = true;
            this.TabControl.Size = new System.Drawing.Size(781, 345);
            this.TabControl.TabIndex = 3;
            this.TabControl.TabStop = false;
            this.TabControl.SelectedIndexChanged += new System.EventHandler(this.TabControl_SelectedIndexChanged);
            this.TabControl.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.TabControl_MouseDoubleClick);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(781, 451);
            this.Controls.Add(this.TabControl);
            this.Controls.Add(this.StatusStrip);
            this.Controls.Add(this.AddressBar);
            this.Controls.Add(this.ToolBar);
            this.Controls.Add(this.MenuStrip);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "浏览器";
            this.MenuStrip.ResumeLayout(false);
            this.MenuStrip.PerformLayout();
            this.ToolBar.ResumeLayout(false);
            this.ToolBar.PerformLayout();
            this.AddressBar.ResumeLayout(false);
            this.AddressBar.PerformLayout();
            this.StatusStrip.ResumeLayout(false);
            this.StatusStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip ToolBar;
        private System.Windows.Forms.ToolStrip AddressBar;
        private System.Windows.Forms.MenuStrip MenuStrip;
        private System.Windows.Forms.StatusStrip StatusStrip;
        private System.Windows.Forms.ToolStripMenuItem MenuFile;
        private System.Windows.Forms.ToolStripMenuItem MenuNew;
        private System.Windows.Forms.ToolStripMenuItem MenuOpen;
        private System.Windows.Forms.ToolStripMenuItem MenuSave;
        private System.Windows.Forms.ToolStripSeparator MenuSeparator1;
        private System.Windows.Forms.ToolStripMenuItem MenuPrint;
        private System.Windows.Forms.ToolStripMenuItem MenuPrintPreview;
        private System.Windows.Forms.ToolStripSeparator MenuSeparator2;
        private System.Windows.Forms.ToolStripMenuItem MenuQuit;
        private System.Windows.Forms.ToolStripMenuItem MenuHelp;
        private System.Windows.Forms.ToolStripMenuItem MenuAbout;
        private System.Windows.Forms.ToolStripSplitButton ToolBack;
        private System.Windows.Forms.ToolStripSplitButton ToolForward;
        private System.Windows.Forms.ToolStripButton ToolStop;
        private System.Windows.Forms.ToolStripButton ToolRefresh;
        private System.Windows.Forms.ToolStripButton ToolHome;
        private System.Windows.Forms.ToolStripLabel ToolLabel;
        private System.Windows.Forms.ToolStripMenuItem MenuSetting;
        private System.Windows.Forms.ToolStripMenuItem MenuView;
        private System.Windows.Forms.ToolStripMenuItem MenuGoto;
        private System.Windows.Forms.ToolStripMenuItem MenuStop;
        private System.Windows.Forms.ToolStripMenuItem MenuRefresh;
        private System.Windows.Forms.ToolStripMenuItem MenuSource;
        private System.Windows.Forms.ToolStripMenuItem MenuBack;
        private System.Windows.Forms.ToolStripButton ToolGoto;
        internal System.Windows.Forms.ToolStripStatusLabel ToolImage;
        internal System.Windows.Forms.ToolStripStatusLabel StatusText;
        internal System.Windows.Forms.ToolStripProgressBar ToolProgressBar;
        internal System.Windows.Forms.ToolStripStatusLabel ToolInternet;
        private System.Windows.Forms.ToolStripComboBox ToolAddress;
        private System.Windows.Forms.ToolStripSplitButton ToolFavorites;
        private System.Windows.Forms.ToolStripMenuItem MenuFavorites;
        private System.Windows.Forms.ToolStripMenuItem MenuForward;
        private System.Windows.Forms.ToolStripMenuItem MenuToolBar;
        private System.Windows.Forms.ToolStripMenuItem MenuAddressBar;
        private System.Windows.Forms.ToolStripMenuItem MenuProperty;
        private System.Windows.Forms.TabControl TabControl;
        private System.Windows.Forms.ToolStripMenuItem MenuTool;
        private System.Windows.Forms.ToolStripMenuItem MenuSnapshot;
        private System.Windows.Forms.ToolStripMenuItem MenuNewTab;
    }
}