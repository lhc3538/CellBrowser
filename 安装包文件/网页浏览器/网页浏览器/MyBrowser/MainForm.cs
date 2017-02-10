using System;
using System.IO;
using System.Net;
using System.Text;
using System.Drawing;
using IWshRuntimeLibrary;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;

namespace MyBrowser
{
    public partial class MainForm : Form
    {
        private List<WebBrowser> WebBrowserList = new List<WebBrowser>(); // WebBrowser数组，记录所有打开的页面信息
        private WebBrowser webBrowser = new WebBrowser(); // 网页截图专用
        private OpenFileDialog openFileDialog = new OpenFileDialog(); // 打开对话框
        private SaveFileDialog saveFileDialog = new SaveFileDialog(); // 保存对话框

        /// <summary>
        /// 初始化窗体
        /// </summary>
        public MainForm()
        {
            InitializeComponent();
            InitHomePage();
            InitFavourites();
            InitAddress();
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized; // 最大化窗体
            this.webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);
        }

        /// <summary>
        /// 初始化首页
        /// </summary>
        private void InitHomePage()
        {
            this.NewWindow(new WebBrowser(), new CancelEventArgs());
            this.ToolHome_Click(null, null);
        }

        /// <summary>
        /// 初始化收藏夹
        /// </summary>
        private void InitFavourites()
        {
            DirectoryInfo directory = new DirectoryInfo("Favourites");
            if (!directory.Exists)
            {
                Directory.CreateDirectory("Favourites");
            }
            foreach (FileInfo file in directory.GetFiles())
            {
                try
                {
                    IWshShell_Class shell = new IWshShell_ClassClass();
                    IWshURLShortcut shortcut = shell.CreateShortcut(file.FullName) as IWshURLShortcut;
                    ToolStripMenuItem menuItem = new ToolStripMenuItem();
                    menuItem.Text = file.Name;
                    menuItem.ToolTipText = shortcut.TargetPath.ToString();
                    menuItem.Click += new System.EventHandler(this.Favourites_Click);
                    this.ToolFavorites.DropDownItems.Add(menuItem);
                }
                catch { }
            }
        }

        /// <summary>
        /// 初始化地址栏的历史记录 
        /// </summary>
        private void InitAddress()
        {
            if (System.IO.File.Exists("History.ini") == false)
            {
                System.IO.File.Create("History.ini");
            }
            else
            {
                StreamReader reader = new StreamReader("History.ini");
                while (reader.Peek() >= 0)
                {
                    this.ToolAddress.Items.Add(reader.ReadLine());
                }
                reader.Close();
            }
        }

        /// <summary>
        /// 快捷方式被单击的事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Favourites_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem menuItem = (ToolStripMenuItem)sender;
            this.WebBrowserList[TabControl.SelectedIndex].Navigate(menuItem.ToolTipText.ToString());
        }

        /// <summary>
        /// 调用WebBrowser事件
        /// </summary>
        private void WebBrowserEventHandler()
        {
            this.WebBrowserList[TabControl.SelectedIndex].NewWindow += new System.ComponentModel.CancelEventHandler(this.NewWindow);
            this.WebBrowserList[TabControl.SelectedIndex].Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.Navigated);
            this.WebBrowserList[TabControl.SelectedIndex].Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.Navigating);
            this.WebBrowserList[TabControl.SelectedIndex].ProgressChanged += new System.Windows.Forms.WebBrowserProgressChangedEventHandler(this.ProgressChanged);
            this.WebBrowserList[TabControl.SelectedIndex].DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.DocumentCompleted);
            //MessageBox.Show(TabControl.SelectedIndex.ToString()+"   "+this.WebBrowserList.Count);
        }

        /// <summary>
        /// 监听回车事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolAddress_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string address = this.ToolAddress.Text.Trim();
                if (address.StartsWith("www.") || address.StartsWith("WWW.")) // 处理输入的URL
                {
                    address = "http://" + address + @"/";
                }
                if (address.StartsWith("http://") || address.StartsWith("ftp://"))
                {
                    this.ToolAddress.Text = address;
                }
                this.TabControl.TabPages[TabControl.SelectedIndex].Text = address;
                this.WebBrowserList[TabControl.SelectedIndex].Dock = DockStyle.Fill;
                this.WebBrowserList[TabControl.SelectedIndex].Navigate(address); // 跳转到URL                
            }
        }

        /// <summary>
        /// WebBrowser_NewWindow
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewWindow(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
            //WebBrowser webBrowser = (WebBrowser)sender;

            WebBrowser webBrowser = new WebBrowser();
            this.WebBrowserList.Add(webBrowser);
            this.TabControl.TabPages.Add("正在载入……");
            this.TabControl.SelectedIndex = this.TabControl.TabPages.Count - 1;
            this.TabControl.TabPages[this.TabControl.TabPages.Count - 1].Controls.Add(WebBrowserList[TabControl.SelectedIndex]);
            this.TabControl.TabPages[this.TabControl.SelectedIndex].Tag = webBrowser;
            this.WebBrowserList[TabControl.SelectedIndex].Dock = DockStyle.Fill;
            this.WebBrowserList[TabControl.SelectedIndex].Navigate(webBrowser.StatusText);
            this.WebBrowserEventHandler();
            this.StatusText.Text = "";
            this.ToolProgressBar.Visible = true;
            this.TabControl.TabPages[TabControl.SelectedIndex].Text = "新建标签";
            this.ToolAddress.Text = webBrowser.StatusText;
            //MessageBox.Show(this.WebBrowserList[TabControl.SelectedIndex].Url.ToString());
        }

        /// <summary>
        /// WebBrowser_Navigated
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            this.TabControl.TabPages[TabControl.SelectedIndex].Text = this.WebBrowserList[TabControl.SelectedIndex].Url.ToString();
            this.StatusText.Text = "正在载入页面";
        }

        /// <summary>
        /// WebBrowser_Navigating
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            this.TabControl.TabPages[TabControl.SelectedIndex].Text = "正在载入……";
            this.StatusText.Text = "";
            this.ToolProgressBar.Visible = true;
        }

        /// <summary>
        /// WebBrowser_ProgressChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ProgressChanged(object sender, WebBrowserProgressChangedEventArgs e)
        {
            this.ToolProgressBar.Value = (int)(((float)e.CurrentProgress / (float)e.MaximumProgress) * 100);
        }

        /// <summary>
        /// WebBrowser_DocumentCompleted
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.TabControl.TabPages[TabControl.SelectedIndex].Text = this.WebBrowserList[TabControl.SelectedIndex].DocumentTitle;
            this.StatusText.Text = "完成";
            this.ToolAddress.Text = this.WebBrowserList[TabControl.SelectedIndex].Url.ToString();
            this.ToolProgressBar.Visible = false;
            this.ToolProgressBar.Value = 0;
            this.ToolState();
            this.FormatAddress();
            this.WebBrowserList[TabControl.SelectedIndex].IsWebBrowserContextMenuEnabled = true; // 屏蔽鼠标右键
            SaveAddress(); // 记录转到的地址
        }

        /// <summary>
        /// 格式化地址栏信息
        /// </summary>
        private void FormatAddress()
        {
            if (this.ToolAddress.Text.StartsWith("file"))
            {
                this.ToolAddress.Text = this.ToolAddress.Text.Replace("file:///", "");
            }
        }

        /// <summary>
        /// 记录转到的地址
        /// </summary>
        private void SaveAddress()
        {
            if (ToolAddress.Text.Equals("about:blank")) return;
            StreamReader reader = new StreamReader("History.ini");
            while (reader.Peek() >= 0)
            {
                if (this.ToolAddress.Text.Equals(reader.ReadLine()))
                {
                    return;
                }
            }
            reader.Close();
            StreamWriter writer = new StreamWriter("History.ini", true);
            writer.WriteLine(this.ToolAddress.Text);
            writer.Flush();
            writer.Close();
        }

        /// <summary>
        /// "后退"与"前进"按钮的状态
        /// </summary>
        private void ToolState()
        {
            if (this.WebBrowserList[TabControl.SelectedIndex].CanGoForward)
            {
                this.ToolForward.Enabled = true;
                this.MenuForward.Enabled = true;
            }
            else
            {
                this.ToolForward.Enabled = false;
                this.MenuForward.Enabled = false;
            }
            if (this.WebBrowserList[TabControl.SelectedIndex].CanGoBack)
            {
                this.ToolBack.Enabled = true;
                this.MenuBack.Enabled = true;
            }
            else
            {
                this.ToolBack.Enabled = false;
                this.MenuBack.Enabled = false;
            }
        }

        /// <summary>
        /// "新建"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuNew_Click(object sender, EventArgs e)
        {
            MainForm newForm = new MainForm();
            newForm.Show();
        }

        /// <summary>
        /// "打开"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuOpen_Click(object sender, EventArgs e)
        {
            openFileDialog.InitialDirectory = Environment.CurrentDirectory + @"\HTML"; // 设定"打开文件"对话框的初始目录
            openFileDialog.Filter = "网页 (*.HTM;*.HTML)|*.HTM;*.HTML|XML 文件(*.XML)|*.XML|Flash 文件(*.SWF)|*.SWF|Office 文件(*.DOC;*.XLS;*.PPT)|*.DOC;*.XLS;*.PPT"; // 设定文件名筛选字符串以方便选取特定格式的文件
            openFileDialog.Title = "请选择特定格式文件"; // 设定"打开"对话框的标题
            openFileDialog.RestoreDirectory = true; // 设定"打开"对话框在关闭前还原目前的目录
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.WebBrowserList[TabControl.SelectedIndex].Navigate(openFileDialog.FileName);
            }
        }

        /// <summary>
        /// "另存为"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuSave_Click(object sender, EventArgs e)
        {
            WebBrowserList[TabControl.SelectedIndex].ShowSaveAsDialog();
        }

        /// <summary>
        /// "打印"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuPrint_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].ShowPrintDialog();
        }

        /// <summary>
        /// "打印预览"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuPrintPreview_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].ShowPrintPreviewDialog();
        }

        /// <summary>
        /// "页面设置"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuSetting_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].ShowPageSetupDialog();
        }

        /// <summary>
        /// "属性"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuProperty_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].ShowPropertiesDialog();
        }

        /// <summary>
        /// "退出"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuQuit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// "工具栏"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuToolBar_Click(object sender, EventArgs e)
        {
            if (this.MenuToolBar.Checked)
            {
                this.MenuToolBar.Checked = false;
                this.ToolBar.Hide();
            }
            else
            {
                this.MenuToolBar.Checked = true;
                this.ToolBar.Show();
            }
        }

        /// <summary>
        /// "地址栏"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuAddressBar_Click(object sender, EventArgs e)
        {
            if (this.MenuAddressBar.Checked)
            {
                this.MenuAddressBar.Checked = false;
                this.AddressBar.Hide();
            }
            else
            {
                this.MenuAddressBar.Checked = true;
                this.AddressBar.Show();
            }
        }

        /// <summary>
        /// "后退"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuBack_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].GoBack();
        }

        /// <summary>
        /// "前进"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuForward_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].GoForward();
        }

        /// <summary>
        /// "停止"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuStop_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].Stop();
        }

        /// <summary>
        /// "刷新"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuRefresh_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].Refresh();
        }

        /// <summary>
        /// "源文件"菜单,将显示在网页的HTML内容显示在一个记事本中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuSource_Click(object sender, EventArgs e)
        {
            if (WebBrowserList[TabControl.SelectedIndex].Url.ToString().StartsWith("http"))
            {
                string fileName = "Source.txt";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(WebBrowserList[TabControl.SelectedIndex].Url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), System.Text.Encoding.GetEncoding("GB2312"));
                FileInfo urlFile = new FileInfo(fileName);
                StreamWriter writer = new StreamWriter(fileName, false);
                writer.WriteLine(reader.ReadToEnd().ToString());
                writer.Flush();
                writer.Close();
                System.Diagnostics.Process.Start("NotePad.exe", fileName);
            }
            else
            {
                MessageBox.Show("这不是一个HTML文件，无法查看其源文件。");
            }
        }

        /// <summary>
        /// "网页截图"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuSnapshot_Click(object sender, EventArgs e)
        {
            saveFileDialog.InitialDirectory = Environment.SpecialFolder.Desktop.ToString(); // 设定"打开文件"对话框的初始目录
            saveFileDialog.Filter = "图片 (*.JPG)|*.JPG"; // 设定文件名筛选字符串以方便选取特定格式的文件
            saveFileDialog.Title = "请选择存储路径"; // 设定"打开"对话框的标题
            saveFileDialog.RestoreDirectory = true; // 设定"打开"对话框在关闭前还原目前的目录
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.webBrowser.Navigate(ToolAddress.Text);
            }
        }

        /// <summary>
        /// 截图处理函数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this.webBrowser.Document == null) return;
            try
            {
                int scrollHeight = this.webBrowser.Document.Body.ScrollRectangle.Height;
                int scrollWidth = this.webBrowser.Document.Body.ScrollRectangle.Width;
                this.webBrowser.Size = new Size(scrollWidth, scrollHeight);
                Bitmap bitmap = new Bitmap(scrollWidth, scrollHeight);
                this.webBrowser.DrawToBitmap(bitmap, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                bitmap.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                bitmap.Dispose();
                MessageBox.Show("保存成功！", "消息提示", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// "关于"菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show("C#基于IE内核的浏览器", "关于");
            return;
        }

        /// <summary>
        /// 工具栏的"添加到收藏夹"菜单,将网页收藏
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuFavorites_Click(object sender, EventArgs e)
        {
            string name = WebBrowserList[TabControl.SelectedIndex].DocumentTitle.ToString();
            string url = this.WebBrowserList[TabControl.SelectedIndex].Url.ToString();
            ToolStripMenuItem menuItem = new ToolStripMenuItem();
            menuItem.Text = name;
            menuItem.ToolTipText = url;
            menuItem.Click += new System.EventHandler(this.Favourites_Click);
            this.ToolFavorites.DropDownItems.Add(menuItem);
            IWshShell_Class shell = new IWshShell_ClassClass();
            IWshURLShortcut shortcut = shell.CreateShortcut("Favourites\\" + name + ".url") as IWshURLShortcut;
            shortcut.TargetPath = url;
            shortcut.Save();
        }

        /// <summary>
        /// 将菜单栏的图标按钮与工具栏的按键连接
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            this.ToolBack.Tag = MenuBack;
            this.ToolForward.Tag = MenuForward;
            this.ToolStop.Tag = MenuStop;
            this.ToolRefresh.Tag = MenuRefresh;
            base.OnLoad(e);
        }

        /// <summary>
        /// 自定义方法,将工具栏的图标跟菜单栏的按键连接起来
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Link_Click(object sender, EventArgs e)
        {
            ToolStripItem toolItem = sender as ToolStripItem;
            if (toolItem != null)
            {
                ToolStripItem menuItem = toolItem.Tag as ToolStripMenuItem;
                if (menuItem != null)
                {
                    menuItem.PerformClick();
                }
            }
        }

        /// <summary>
        /// "主页"按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolHome_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].GoHome();
        }

        /// <summary>
        /// "转到"按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolGoto_Click(object sender, EventArgs e)
        {
            this.WebBrowserList[TabControl.SelectedIndex].Navigate(this.ToolAddress.Text);
        }

        /// <summary>
        /// 改变浏览器大小时,自动改变控件尺寸
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddressBar_SizeChanged(object sender, EventArgs e)
        {
            this.ToolAddress.Size = new Size(this.AddressBar.Width - 120, this.ToolAddress.Height);
        }

        /// <summary>
        /// 双击关闭选中的面板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabControl_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            TabControl TabControl = (TabControl)sender;
            Point point = new Point(e.X, e.Y);
            if (TabControl.TabCount == 1)
            {
                return;
            }
            for (int i = 0; i < TabControl.TabCount; i++)
            {
                Rectangle recTab = TabControl.GetTabRect(i);
                if (recTab.Contains(point))
                {
                    TabPage seltab = this.TabControl.SelectedTab;
                    int seltabindex = this.TabControl.SelectedIndex;
                    this.WebBrowserList.Remove((WebBrowser)seltab.Tag);
                    TabControl.Controls.Remove(seltab);
                }
            }
        }

        private void MenuNewTab_Click(object sender, EventArgs e)
        {
            WebBrowser webBrowser = new WebBrowser();
            this.WebBrowserList.Add(webBrowser);
            this.TabControl.TabPages.Add("正在载入……");
            this.TabControl.SelectedIndex = this.TabControl.TabPages.Count - 1;
            this.TabControl.TabPages[this.TabControl.TabPages.Count - 1].Controls.Add(WebBrowserList[TabControl.SelectedIndex]);
            this.TabControl.TabPages[this.TabControl.SelectedIndex].Tag = webBrowser;
            this.WebBrowserList[TabControl.SelectedIndex].Dock = DockStyle.Fill;
            this.WebBrowserList[TabControl.SelectedIndex].Navigate(webBrowser.StatusText);
            this.WebBrowserEventHandler();
            this.StatusText.Text = "";
            this.ToolProgressBar.Visible = true;
            this.TabControl.TabPages[TabControl.SelectedIndex].Text = "新建标签";
            this.ToolAddress.Text = webBrowser.StatusText;
        }

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)this.TabControl.TabPages[this.TabControl.SelectedIndex].Tag;
            if (webBrowser!=null && webBrowser.ReadyState == WebBrowserReadyState.Complete)
            {
                this.ToolAddress.Text = webBrowser.Url.ToString();
            }
        }
    }
}