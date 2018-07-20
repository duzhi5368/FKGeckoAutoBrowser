using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Threading;
using FKWebBrowser.Forms;
using FKWebBrowser.Source;
using Gecko;
using Gecko.DOM;
using Gecko.Plugins;
using Gecko.JQuery;
using System.Diagnostics;
using System.Collections;
using System.Linq;

namespace FKWebBrowser
{
    [System.Security.Permissions.PermissionSet(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]

    public partial class Form_Main : Form
    {
        #region 类静态变量
        // 当前打开的Tab选项卡
        private TabPage s_CurrentTab = null;
        // 临时Html元素对象
        private GeckoHtmlElement s_HtmlElement;
        // 主浏览器
        private WebBrowser s_MainWebBrowser;
        // Hook对象
        private MouseKeyboardLibrary.MouseHook s_MouseHook = new MouseKeyboardLibrary.MouseHook();
        private MouseKeyboardLibrary.KeyboardHook s_KeyboardHook = new MouseKeyboardLibrary.KeyboardHook();
        // 上次Hook时间
        private int s_LastTimeRecorded = 0;
        // 当前脚本窗口是否显示
        private bool s_IsScriptWindowShow = true;
        // 当前脚本文件名
        private string s_StrLastScriptFileName = "";
        // 当前软件版本号
        private string s_CUR_APP_VERSION = "1.1.8";
        // 脚本是否停止
        private bool s_IsScriptStop = false;
        // 弹出菜单时的鼠标位置
        private Point s_tagRightClickPoint;
        // 
        private FKConsoleWrapper s_Wrapper;

        // 是否页面加载完毕
        private bool s_IsPageReady { get; set; }
        #endregion
        public Form_Main(FKConsoleWrapper p)
        {
            InitializeComponent();
            s_Wrapper = p;
        }

        // 进程的初始化
        private void AppInit()
        {
            // 初始化日志
            s_Wrapper.GetLog().SetForm(this);
            // 初始化JS脚本
            InitJS();
            // 初始化键盘鼠标Hook
            InitMouseAndKeyboardHook();
            // 旅游者模式关闭脚本窗口
            if (s_Wrapper.GetUserMode() == FKConsoleWrapper.ENUM_UserMode.eUserMode_Guest)
            {
              ToggleScriptWindowVisible();        // 默认隐藏
              tabControlCode.Visible = false;
              tabControlCode.Enabled = false;
              FKLog("当前用户模式：游客模式");
            }
            else
            {
                FKLog("当前用户模式：开发者模式");
            }
            // 启动Gecko内核
            try {
                GeckoWebBrowser.UseCustomPrompt();
                Xpcom.Initialize("XULRunner");
            }
            catch (Exception e)
            {
                FKLog("启动XULRunner失败:" + e.ToString());
            }
            // 激活全部插件
            ActiveAllPlugins();
            // 一些默认设置
            SetDefaultPreferences();
            // 初始化定时内存清理
            CreateXULMemoryGCTimer();
        }

        private void FKDebugLog(string str)
        {
            //if (s_Wrapper.GetUserMode() == FKConsoleWrapper.ENUM_UserMode.eUserMode_Guest ||
            //    s_Wrapper.GetUserMode() == FKConsoleWrapper.ENUM_UserMode.eUserMode_Developer)
            //{
                Console.WriteLine("[debug] " + str);
            //}
        }

        private void ActiveAllPlugins()
        {
            var vecPlugins = PluginHost.GetPluginTags();
            foreach (var plugin in vecPlugins)
            {
                plugin.Blocklisted = false;
                plugin.Disabled = false;
                //FKLog(string.Format("{0} 插件：{1}..{2},路径：{3}", plugin.Disabled ? "禁止" : "允许", plugin.Name, plugin.Version, plugin.Fullpath));
            }

            GeckoPreferences.Default["extensions.blocklist.enabled"] = false;
        }
        private void SetDefaultPreferences()
        {
            try {
                GeckoPreferences.User["devtools.debugger.remote-enabled"] = true;
                GeckoPreferences.User["browser.xul.error_pages.enabled"] = false;
                GeckoPreferences.User["browser.download.manager.showAlertOnComplete"] = true;
                GeckoPreferences.User["security.warn_viewing_mixed"] = false;
                GeckoPreferences.User["privacy.popups.showBrowserMessage"] = false;

                GeckoPreferences.User["browser.download.useDownloadDir"] = true;
                GeckoPreferences.User["browser.download.folderList"] = 0;
                GeckoPreferences.User["browser.download.manager.showAlertOnComplete"] = true;
                GeckoPreferences.User["browser.download.manager.showAlertInterval"] = 2000;
                GeckoPreferences.User["browser.download.manager.retention"] = 2;
                GeckoPreferences.User["browser.download.manager.showWhenStarting"] = true;
                GeckoPreferences.User["browser.download.manager.useWindows]"] = true;
                GeckoPreferences.User["browser.download.manager.closeWhenDone"] = true;
                GeckoPreferences.User["browser.download.manager.openDelay"] = 0;
                GeckoPreferences.User["browser.download.manager.focusWhenStarting"] = false;
                GeckoPreferences.User["browser.download.manager.flashCount"] = 2;
            }
            catch { }
        }
        private void CreateXULMemoryGCTimer()
        {
            try {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                if (timer != null)
                {
                    timer.Interval = (2 * 60 * 1000);
                    timer.Tick += new EventHandler(XULMemoryGVTimeTicker);
                    timer.Start();
                }
            }
            catch { }
        }
        private void XULMemoryGVTimeTicker(object send, EventArgs e)
        {
            try {
                var memoryService = Xpcom.GetService<nsIMemory>("@mozilla.org/xpcom/memory-service;1");
                memoryService.HeapMinimize(false);
            }
            catch { }
        }

        // 窗口的初始化
        private void OnWindowInit()
        {
            try {
                // 截获按键消息
                this.KeyUp += new System.Windows.Forms.KeyEventHandler(MyHotKeyEvent);

                // 语言初始化
                UpdateLanguage("zh-CN");

                // 创建一个默认空窗口
                CreateNewTab();
            }
            catch (Exception e)
            {
                FKLog("窗口初始化失败: " + e.ToString());
            }
            // 初始化完成
            FKLog(Language.LanguageRes.Log_InitDone);
        }

        #region 更新控件语言
        // 更新控件语言
        public void UpdateLanguage(string strLanguage)
        {
            try {
                Language.LanguageRes.Culture = CultureInfo.CreateSpecificCulture(strLanguage);
                System.Reflection.Assembly.GetExecutingAssembly();

                // 窗口标题修改
                this.Text = Language.LanguageRes.WndTiltle_AppName + " - " + Language.LanguageRes.WndTitle_AppVersionDesc + " - " + s_CUR_APP_VERSION;

                // 控件名修改
                InitUIText();

                // 重构右键菜单
                InitContextMenu();
            } catch (Exception e)
            {
                throw new Exception("更新控件状态失败: " + e.ToString());
            }
        }

        // 重置控件名
        private void InitUIText()
        {
            scriptToolStripMenuItem.Text = Language.LanguageRes.UI_ScriptMenu;
            newToolStripMenuItem.Text = Language.LanguageRes.UI_NewScriptMenu;
            openToolStripMenuItem.Text = Language.LanguageRes.UI_OpenScriptMenu;
            saveToolStripMenuItem.Text = Language.LanguageRes.UI_SaveScriptMenu;
            saveAsToolStripMenuItem.Text = Language.LanguageRes.UI_SaveAsScriptMenu;

            accountToolStripMenuItem.Text = Language.LanguageRes.UI_AccountMenu;
            loginToolStripMenuItem.Text = Language.LanguageRes.UI_LoginAccountMenu;
            registerToolStripMenuItem.Text = Language.LanguageRes.UI_RegisterAccountMenu;
            logoutToolStripMenuItem.Text = Language.LanguageRes.UI_LogoutAccountMenu;

            viewToolStripMenuItem.Text = Language.LanguageRes.UI_ViewMenu;
            scriptWindowToolStripMenuItem.Text = Language.LanguageRes.UI_ScriptWindowMenu;

            settingToolStripMenuItem.Text = Language.LanguageRes.UI_SettingMenu;
            aboutToolStripMenuItem.Text = Language.LanguageRes.UI_SettingAboutMenu;
            hideBrowserToolStripMenuItem.Text = Language.LanguageRes.UI_SettingHideAppMenu;
            languageToolStripMenuItem.Text = Language.LanguageRes.UI_SettingLanguageMenu;
            zhCNToolStripMenuItem.Text = Language.LanguageRes.UI_SettingLangZH_CNMenu;
            enUSToolStripMenuItem.Text = Language.LanguageRes.UI_SettingLangEN_USMenu;

            toolStripBack.Text = Language.LanguageRes.UI_BackMenu;
            toolStripNext.Text = Language.LanguageRes.UI_NextMenu;
            toolStripReload.Text = Language.LanguageRes.UI_RefreshMenu;
            btnNewTab.Text = Language.LanguageRes.UI_NewTabMenu;
            btnCloseTab.Text = Language.LanguageRes.UI_CloseTabMenu;
            btnCloseAllTab.Text = Language.LanguageRes.UI_CloseAllTabMenu;

            btnGo.Text = Language.LanguageRes.UI_GoBtn;

            toolStripRunScript.Text = Language.LanguageRes.UI_RunScriptBtn;
            toolStripRecordScript.Text = Language.LanguageRes.UI_RecordBtn;
            toolStripNewScript.Text = Language.LanguageRes.UI_NewScriptMenu;
            toolStripOpenScript.Text = Language.LanguageRes.UI_OpenScriptMenu;
            toolStripSaveScript.Text = Language.LanguageRes.UI_SaveScriptMenu;
            toolStripSaveScriptAs.Text = Language.LanguageRes.UI_SaveAsScriptMenu;
            toolStripClearScript.Text = Language.LanguageRes.UI_ClearScriptBtn;

            tabPageLog.Text = Language.LanguageRes.UI_LogTabPage;
            tagPageAutoScript.Text = Language.LanguageRes.UI_ScriptTabPage;
            tabPageMP.Text = Language.LanguageRes.UI_MPTabPage;
            tabPageTPP.Text = Language.LanguageRes.UI_TPPTabPage;
        }

        // 重置右键菜单
        private void InitContextMenu()
        {
            contextMenuBrowser.Items.Clear();
            if (s_Wrapper.GetUserMode() != FKConsoleWrapper.ENUM_UserMode.eUserMode_Developer)
                return;

            var baseItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Base);
            var goItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Go);
            var sleepItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Sleep);
            var createfolderItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Base_CreateFolder);
            var removefolderItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Base_RemoveFolder);
            var createfileItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Base_CreateFile);
            var getfilesItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Base_GetFiles);
            var getfoldersItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Base_GetFolders);
            var removefileItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Base_RemoveFile);
            var openDirInExploreItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Base_OpenExplore);
            goItem.Click += MenuItemClick;
            sleepItem.Click += MenuItemClick;
            createfolderItem.Click += MenuItemClick;
            removefolderItem.Click += MenuItemClick;
            createfileItem.Click += MenuItemClick;
            getfilesItem.Click += MenuItemClick;
            getfoldersItem.Click += MenuItemClick;
            removefileItem.Click += MenuItemClick;
            openDirInExploreItem.Click += MenuItemClick;
            baseItem.DropDownItems.Add(goItem);
            baseItem.DropDownItems.Add(sleepItem);
            baseItem.DropDownItems.Add(createfolderItem);
            baseItem.DropDownItems.Add(removefolderItem);
            baseItem.DropDownItems.Add(createfileItem);
            baseItem.DropDownItems.Add(getfilesItem);
            baseItem.DropDownItems.Add(getfoldersItem);
            baseItem.DropDownItems.Add(removefileItem);
            baseItem.DropDownItems.Add(openDirInExploreItem);
            contextMenuBrowser.Items.Add(baseItem);

            var advancedItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Advanced);
            var runCommandItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Advanced_RunCommand);
            var imageToTextItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Advanced_ImageToText);
            var takesnapshotItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Advanced_TakeSnapShot);
            var sendEmailItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Advanced_SendEmail);
            var readExcelItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Advanced_ReadExcel);
            var writeExcelItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Advanced_WriteExcel);
            runCommandItem.Click += MenuItemClick;
            imageToTextItem.Click += MenuItemClick;
            takesnapshotItem.Click += MenuItemClick;
            sendEmailItem.Click += MenuItemClick;
            readExcelItem.Click += MenuItemClick;
            writeExcelItem.Click += MenuItemClick;
            advancedItem.DropDownItems.Add(runCommandItem);
            advancedItem.DropDownItems.Add(imageToTextItem);
            advancedItem.DropDownItems.Add(takesnapshotItem);
            advancedItem.DropDownItems.Add(sendEmailItem);
            advancedItem.DropDownItems.Add(readExcelItem);
            advancedItem.DropDownItems.Add(writeExcelItem);
            contextMenuBrowser.Items.Add(advancedItem);

            var extractItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Extract);
            var attributeItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Extract_Attribute);
            var htmlItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Extract_Html);
            var srcItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Extract_Src);
            var textItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Extract_Text);
            var urlItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Extract_Url);
            attributeItem.Click += MenuItemClick;
            htmlItem.Click += MenuItemClick;
            srcItem.Click += MenuItemClick;
            textItem.Click += MenuItemClick;
            urlItem.Click += MenuItemClick;
            extractItem.DropDownItems.Add(attributeItem);
            extractItem.DropDownItems.Add(htmlItem);
            extractItem.DropDownItems.Add(srcItem);
            extractItem.DropDownItems.Add(textItem);
            extractItem.DropDownItems.Add(urlItem);
            contextMenuBrowser.Items.Add(extractItem);

            var fillItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Fill);
            var textboxItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Fill_Textbox);
            var dropdownItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Fill_Dropdown);
            var iFrameItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Fill_iFrame);
            var clickElementItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Fill_ClickElement);
            textboxItem.Click += MenuItemClick;
            dropdownItem.Click += MenuItemClick;
            iFrameItem.Click += MenuItemClick;
            clickElementItem.Click += MenuItemClick;
            fillItem.DropDownItems.Add(textboxItem);
            fillItem.DropDownItems.Add(dropdownItem);
            fillItem.DropDownItems.Add(iFrameItem);
            fillItem.DropDownItems.Add(clickElementItem);
            contextMenuBrowser.Items.Add(fillItem);

            var mouseItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Mouse);
            var currentMouseItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Mouse_GetCurrentMouse);
            var mouseMoveItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Mouse_MouseMove);
            var mouseDownItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Mouse_MouseDown);
            var mouseUpItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Mouse_MouseUp);
            var mouseClickItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Mouse_MouseClick);
            var mouseDoubleClickItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Mouse_MouseDoubleClick);
            var mouseWheelItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Mouse_MouseWheel);
            currentMouseItem.Click += MenuItemClick;
            mouseMoveItem.Click += MenuItemClick;
            mouseDownItem.Click += MenuItemClick;
            mouseUpItem.Click += MenuItemClick;
            mouseClickItem.Click += MenuItemClick;
            mouseDoubleClickItem.Click += MenuItemClick;
            mouseWheelItem.Click += MenuItemClick;
            mouseItem.DropDownItems.Add(currentMouseItem);
            mouseItem.DropDownItems.Add(mouseMoveItem);
            mouseItem.DropDownItems.Add(mouseDownItem);
            mouseItem.DropDownItems.Add(mouseUpItem);
            mouseItem.DropDownItems.Add(mouseClickItem);
            mouseItem.DropDownItems.Add(mouseDoubleClickItem);
            mouseItem.DropDownItems.Add(mouseWheelItem);
            contextMenuBrowser.Items.Add(mouseItem);

            var keyboardItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Keyboard);
            var keyDownItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Keyboard_KeyDown);
            var keyUpItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Keyboard_KeyUp);
            keyDownItem.Click += MenuItemClick;
            keyUpItem.Click += MenuItemClick;
            keyboardItem.DropDownItems.Add(keyDownItem);
            keyboardItem.DropDownItems.Add(keyUpItem);
            contextMenuBrowser.Items.Add(keyboardItem);

            var databaseItem = new ToolStripMenuItem(Language.LanguageRes.UI_ContextMenu_Database);
            contextMenuBrowser.Items.Add(databaseItem);
        }
        #endregion

        // 添加元素回调
        void MenuItemClick(Object sender, EventArgs e)
        {
            string xpath = string.Empty;
            xpath = s_HtmlElement.GetAttribute("id");
            contextMenuBrowser.Hide();

            if (string.IsNullOrEmpty(xpath))
            {
                // 如果有元素参数，则将元素转为 XPath 格式
                xpath = GetXpath(s_HtmlElement);
            }

            string item = sender.ToString();
            if (item == Language.LanguageRes.UI_ContextMenu_Mouse_GetCurrentMouse)
            {
                string strInfo = "X : " + s_tagRightClickPoint.X.ToString() + ", Y : " + s_tagRightClickPoint.Y.ToString();
                Prompt.ShowDialog(Language.LanguageRes.Msgbox_CurPos,
                    Language.LanguageRes.Msgbox_Message, strInfo, false);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Mouse_MouseMove)
            {
                var x = s_tagRightClickPoint.X.ToString();
                var y = s_tagRightClickPoint.Y.ToString();
                tbxCode.AppendText(Language.LanguageRes.Code_CommitMouseMove + Environment.NewLine);
                tbxCode.AppendText("MouseMove(" + x + ", " + y + ", true, 0);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Mouse_MouseDown)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitMouseDown + Environment.NewLine);
                tbxCode.AppendText("MouseDown('Left', 0);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Mouse_MouseUp)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitMouseUp + Environment.NewLine);
                tbxCode.AppendText("MouseUp('Left', 0);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Mouse_MouseClick)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitMouseClick + Environment.NewLine);
                tbxCode.AppendText("MouseClick('Left', 0);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Mouse_MouseDoubleClick)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitMouseDoubleClick + Environment.NewLine);
                tbxCode.AppendText("MouseDoubleClick('Left', 0);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Mouse_MouseWheel)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitMouseWheel + Environment.NewLine);
                tbxCode.AppendText("MouseWheel(-15, 0);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Keyboard_KeyDown)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitKeyDown + Environment.NewLine);
                tbxCode.AppendText("KeyDown('A', 0);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Keyboard_KeyUp)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitKeyUp + Environment.NewLine);
                tbxCode.AppendText("KeyUp('A', 0);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Go)
            {
                if (MessageBox.Show(Language.LanguageRes.Msgbox_ConfirmGoWebsite,
                    Language.LanguageRes.Msgbox_Message, MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    string address = tbxAddress.Text;
                    string promptValue = Prompt.ShowDialog(Language.LanguageRes.Msgbox_Website,
                        Language.LanguageRes.Msgbox_Message, address, false);
                    if (!string.IsNullOrEmpty(promptValue))
                    {
                        tbxCode.AppendText(Language.LanguageRes.Code_CommitGo + Environment.NewLine);
                        tbxCode.AppendText("go(\"" + promptValue + "\");\n");
                    }
                }
                else
                {
                    if (xpath != "top")
                    {
                        tbxCode.AppendText(Language.LanguageRes.Code_CommitGo + Environment.NewLine);
                        tbxCode.AppendText("go(\"" + xpath + "\");\n");
                    }
                }
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Sleep)
            {
                string promptValue = Prompt.ShowDialog(Language.LanguageRes.Msgbox_SleepInfo,
                    Language.LanguageRes.Msgbox_Message, "1", false);
                if (MessageBox.Show(Language.LanguageRes.Msgbox_ConfirmSleep,
                    Language.LanguageRes.Msgbox_Message, MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    tbxCode.AppendText(Language.LanguageRes.Code_CommitSleepTrue + Environment.NewLine);
                    tbxCode.AppendText("sleep(" + promptValue + ",true);\n");
                }
                else
                {
                    tbxCode.AppendText(Language.LanguageRes.Code_CommitSleepFalse + Environment.NewLine);
                    tbxCode.AppendText("sleep(" + promptValue + ",false);\n");
                }
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Extract_Attribute)
            {
                string promptValue = Prompt.ShowDialog(Language.LanguageRes.Msgbox_Attribute, Language.LanguageRes.Msgbox_Message, "", false);
                if (!string.IsNullOrEmpty(promptValue))
                {
                    tbxCode.AppendText(Language.LanguageRes.Code_CommitExtractAttribute + Environment.NewLine);
                    tbxCode.AppendText("var attribute = extract(\"" + xpath + "\", \"" + promptValue + "\");" + Environment.NewLine);
                    tbxCode.AppendText("log(attribute);" + Environment.NewLine);
                }
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Extract_Html)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitExtractAttribute + Environment.NewLine);
                tbxCode.AppendText("var html = extract(\"" + xpath + "\", \"html\");" + Environment.NewLine);
                tbxCode.AppendText("log(html);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Extract_Src)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitExtractAttribute + Environment.NewLine);
                tbxCode.AppendText("var src = extract(\"" + xpath + "\", \"src\");" + Environment.NewLine);
                tbxCode.AppendText("log(src);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Extract_Text)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitExtractAttribute + Environment.NewLine);
                tbxCode.AppendText("var text = extract(\"" + xpath + "\", \"text\");" + Environment.NewLine);
                tbxCode.AppendText("log(text);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Extract_Url)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitExtractAttribute + Environment.NewLine);
                tbxCode.AppendText("var url = extract(\"" + xpath + "\", \"href\");" + Environment.NewLine);
                tbxCode.AppendText("log(url);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Fill_Textbox)
            {
                string promptValue = Prompt.ShowDialog(Language.LanguageRes.Msgbox_Textbox, Language.LanguageRes.Msgbox_Message, "", false);
                if (!string.IsNullOrEmpty(promptValue))
                {
                    tbxCode.AppendText(Language.LanguageRes.Code_CommitAutoFillTextbox + Environment.NewLine);
                    tbxCode.AppendText("fill(\"" + xpath + "\", \"" + promptValue + "\");" + Environment.NewLine);
                }
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Fill_Dropdown)
            {
                string promptValue = Prompt.ShowDialog(Language.LanguageRes.Msgbox_Dropdown, Language.LanguageRes.Msgbox_Message, "", false);
                if (!string.IsNullOrEmpty(promptValue))
                {
                    tbxCode.AppendText(Language.LanguageRes.Code_CommitAutoFillDropdown + Environment.NewLine);
                    tbxCode.AppendText("filldropdown(\"" + xpath + "\", \"" + promptValue + "\");" + Environment.NewLine);
                }
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Fill_iFrame)
            {
                string promptValue = Prompt.ShowDialog(Language.LanguageRes.Msgbox_iFrame, Language.LanguageRes.Msgbox_Message, "", false);
                if (!string.IsNullOrEmpty(promptValue))
                {
                    tbxCode.AppendText(Language.LanguageRes.Code_CommitAutoFilliFrame + Environment.NewLine);
                    tbxCode.AppendText("filliframe(\"title\", \"" + promptValue + "\");" + Environment.NewLine);
                }
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Fill_ClickElement)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitAutoClinkElement + Environment.NewLine);
                tbxCode.AppendText("click(\"" + xpath + "\");" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Base_CreateFolder)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitCreateFolder + Environment.NewLine);
                tbxCode.AppendText("createfolder('testPath');" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Base_RemoveFolder)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitRemoveFolder + Environment.NewLine);
                tbxCode.AppendText("removefolder('testPath');" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Base_CreateFile)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitCreateFile + Environment.NewLine);
                tbxCode.AppendText("save('this is some content 这是一些测试文本', './test.txt', 'true');" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Base_GetFiles)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitGetFiles + Environment.NewLine);
                tbxCode.AppendText("var files= getfiles('testPath');" + Environment.NewLine);
                tbxCode.AppendText("log(files);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Base_GetFolders)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitGetFolders + Environment.NewLine);
                tbxCode.AppendText("var folders = getfolders('testPath');" + Environment.NewLine);
                tbxCode.AppendText("log(folders);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Base_RemoveFile)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitRemoveFile + Environment.NewLine);
                tbxCode.AppendText("remove('./test.txt');" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Base_OpenExplore)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitOpenExplore + Environment.NewLine);
                tbxCode.AppendText("explorer('testPath');" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Advanced_RunCommand)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitRunCommand + Environment.NewLine);
                tbxCode.AppendText("runcommand('notepad', 'params.txt');" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Advanced_ImageToText)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitImageToText + Environment.NewLine);
                tbxCode.AppendText("var text = imageToText('" + xpath + "', 'eng');" + Environment.NewLine);
                tbxCode.AppendText("log(text);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Advanced_TakeSnapShot)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitSnapShot + Environment.NewLine);
                tbxCode.AppendText("var location = getCurrentPath() + '\\\\ScreenSnapShot.png';" + Environment.NewLine);
                tbxCode.AppendText("takesnapshot(location);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Advanced_SendEmail)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitSendEmail + Environment.NewLine);
                tbxCode.AppendText("var email = sendEmail('name', 'email', 'subject', 'content');" + Environment.NewLine);
                tbxCode.AppendText("log(email);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Advanced_ReadExcel)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitReadExcel + Environment.NewLine);
                tbxCode.AppendText("var readItem = readCellExcel('filePath', 'sheetname', 'row', 'column');" + Environment.NewLine);
                tbxCode.AppendText("log(readItem);" + Environment.NewLine);
            }
            else if (item == Language.LanguageRes.UI_ContextMenu_Advanced_WriteExcel)
            {
                tbxCode.AppendText(Language.LanguageRes.Code_CommitWriteExcel + Environment.NewLine);
                tbxCode.AppendText("writeCellExcel('filePath', 'sheetname', 'A1', 'value');" + Environment.NewLine);
            }
        }

        // 将元素转为XPath格式
        private string GetXpath(GeckoHtmlElement htmlEle)
        {
            string strXPath = "";
            try
            {
                while (htmlEle != null)
                {
                    int ind = GetXpathIndex(htmlEle);
                    if (ind > 1)
                        strXPath = "/" + htmlEle.TagName.ToLower() + "[" + ind + "]" + strXPath;
                    else
                        strXPath = "/" + htmlEle.TagName.ToLower() + strXPath;

                    htmlEle = htmlEle.Parent;
                }
            }
            catch (Exception e)
            {
                FKLog("获取元素XPath失败: " + e.ToString());
            }
            return strXPath;
        }

        // 获取元素 XPath 索引编号
        private int GetXpathIndex(GeckoHtmlElement htmlEle)
        {
            if (htmlEle.Parent == null)
                return 0;

            try {
                int nIndex = 0;
                int nIndexEle = 0;
                string strTagName = htmlEle.TagName;
                GeckoNodeCollection listChildNodes = htmlEle.Parent.ChildNodes;
                foreach (GeckoNode it in listChildNodes)
                {
                    if (it.NodeName == strTagName)
                    {
                        nIndex++;
                        if (it.TextContent == htmlEle.TextContent)
                            nIndexEle = nIndex;
                    }
                }
                if (nIndex > 1)
                    return nIndexEle;
            }
            catch (Exception e)
            {
                FKLog("获取元素XPath索引号失败: " + e.ToString());
            }
            return 0;
        }

        // 获取当前路径
        public string GetCurrentPath()
        {
            string result = Application.StartupPath;
            if (result == "")
                result = Application.UserAppDataPath;
            return result;
        }

        // 设置HTTP代理
        public void SetHttpProxy(string strProxyIP, string strProxyPort)
        {
            if (strProxyIP == "" && strProxyPort == "")
            {
                FKLog(Language.LanguageRes.Log_NoProxy);
                GeckoPreferences.Default["network.proxy.type"] = 0;
                return;
            }

            GeckoPreferences.Default["network.proxy.type"] = 1;
            GeckoPreferences.Default["network.proxy.http"] = strProxyIP;
            GeckoPreferences.Default["network.proxy.http_port"] = Int32.Parse(strProxyPort);
            GeckoPreferences.Default["network.proxy.ssl"] = strProxyIP;
            GeckoPreferences.Default["network.proxy.ssl_port"] = Int32.Parse(strProxyPort);

            FKLog(Language.LanguageRes.Log_UseHttpProxyIP + strProxyIP +
                Language.LanguageRes.Log_UseHttpProxyPort + strProxyPort);
        }

        // 保存图片
        private bool SaveImage(string xpath, string location)
        {
            bool result = false;
            GeckoWebBrowser wbBrowser = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (wbBrowser == null)
                return result;

            GeckoImageElement element = null;
            if (xpath.StartsWith("/"))
                element = (GeckoImageElement)GetElementByXpath(wbBrowser.Document, xpath);
            else
                element = (GeckoImageElement)wbBrowser.Document.GetElementById(xpath);
            if (element == null)
                return false;
            /*
            GeckoSelection selection = wbBrowser.Window.Selection;
            selection.SelectAllChildren(element);
            wbBrowser.CopyImageContents();
            if (Clipboard.ContainsImage())
            {
                Image img = Clipboard.GetImage();
                img.Save(location, System.Drawing.Imaging.ImageFormat.Jpeg);
                result = true;
            }
            */
            Image img = (Bitmap)Image.FromStream( new MemoryStream(Gecko.Utils.SaveImageElement.ConvertGeckoImageElementToPng(wbBrowser,
    element, 0, 0, element.OffsetWidth, element.OffsetHeight)));
            img.Save(location, System.Drawing.Imaging.ImageFormat.Png);
            result = true;
            return result;
        }

        // 读取一个文件
        public string Read(string path)
        {
            string result = "";
            string[] list = System.IO.File.ReadAllLines(path);
            foreach (string l in list)
            {
                result += l + "\n";
            }
            return result;
        }

        // 在Log窗口添加日志
        public void FKLog(string strText)
        {
            if (s_Wrapper.GetLog() != null)
                s_Wrapper.GetLog().AddInfo(strText);
            else
                this.tbxLog.Text += string.Format("{0} | {1}\r\n", DateTime.Now.ToString(), strText);
        }

        public void AddInfoToForm(string strText)
        {
            this.tbxLog.Invoke(new Action<string>((m) =>
            {
                this.tbxLog.Text += (m + "\r\n");
            }), strText);
        }

        // 设置脚本
        public void SetTabCodeText_OtherThread(string strText)
        {
            this.tbxCode.Invoke(new Action<string>((m) =>
            {
                this.tbxCode.Text = m;
            }), strText);
        }

        public void FKLog_OtherThread(string strInfo)
        {
            if (s_Wrapper.GetLog() != null)
                s_Wrapper.GetLog().AddInfo(strInfo);
            else
            {
                this.tbxLog.Invoke(new Action<string>((m) =>
                {
                    this.tbxLog.Text += string.Format("{0} | {1}\r\n", DateTime.Now.ToString(), m);
                }), strInfo);
            }
        }

        // 自己的按键消息处理函数
        private void MyHotKeyEvent(object sender, KeyEventArgs e)
        {
            if (s_Wrapper.GetUserMode() != FKConsoleWrapper.ENUM_UserMode.eUserMode_Developer)
                return;

            switch (e.KeyValue)
            {
                case 116: // f5
                    toolStripRunScript_Click(this, null);
                    break;
                case 117: // f6
                    toolStripRecordScript_Click(this, null);
                    break;
                default:
                    break;
            }
        }

        // 获取当前鼠标 X 坐标点
        public string GetCurrentMouseX()
        {
            return Form_Main.MousePosition.X.ToString();
        }

        // 获取当前鼠标 Y 坐标点
        public string GetCurrentMouseY()
        {
            return Form_Main.MousePosition.Y.ToString();
        }

        // 初始化键盘和鼠标Hook
        private void InitMouseAndKeyboardHook()
        {
            if (s_Wrapper.GetUserMode() != FKConsoleWrapper.ENUM_UserMode.eUserMode_Developer)
                return;

            try
            {
                s_MouseHook.MouseMove += new MouseEventHandler(MouseHook_MouseMove);
                s_MouseHook.MouseDown += new MouseEventHandler(MouseHook_MouseDown);
                s_MouseHook.MouseUp += new MouseEventHandler(MouseHook_MouseUp);
                s_MouseHook.MouseWheel += new MouseEventHandler(MouseHook_MouseWheel);

                s_KeyboardHook.KeyDown += new KeyEventHandler(KeyboardHook_KeyDown);
                s_KeyboardHook.KeyUp += new KeyEventHandler(KeyboardHook_KeyUp);
            }
            catch (Exception e) {
                FKLog("初始化键盘和鼠标Hook失败:" + e.ToString());
            }
        }
        #region Hook回调
        void MouseHook_MouseMove(object sender, MouseEventArgs e)
        {
            tbxCode.AppendText("MouseMove(" + e.X + "," + e.Y + ",true, " + (Environment.TickCount - s_LastTimeRecorded) + ");" + Environment.NewLine);
            s_LastTimeRecorded = Environment.TickCount;
        }

        void MouseHook_MouseDown(object sender, MouseEventArgs e)
        {
            tbxCode.AppendText("MouseDown('" + e.Button.ToString() + "', " + (Environment.TickCount - s_LastTimeRecorded) + ");" + Environment.NewLine);
            s_LastTimeRecorded = Environment.TickCount;
        }

        void MouseHook_MouseUp(object sender, MouseEventArgs e)
        {
            tbxCode.AppendText("MouseUp('" + e.Button.ToString() + "', " + (Environment.TickCount - s_LastTimeRecorded) + ");" + Environment.NewLine);
            s_LastTimeRecorded = Environment.TickCount;
        }

        void MouseHook_MouseWheel(object sender, MouseEventArgs e)
        {
            tbxCode.AppendText("MouseWheel(" + e.Delta + ", " + (Environment.TickCount - s_LastTimeRecorded) + ");" + Environment.NewLine);
            s_LastTimeRecorded = Environment.TickCount;
        }

        void KeyboardHook_KeyDown(object sender, KeyEventArgs e)
        {
            tbxCode.AppendText("KeyDown('" + e.KeyCode + "', " + (Environment.TickCount - s_LastTimeRecorded) + ");" + Environment.NewLine);
            s_LastTimeRecorded = Environment.TickCount;
        }

        void KeyboardHook_KeyUp(object sender, KeyEventArgs e)
        {
            tbxCode.AppendText("KeyUp('" + e.KeyCode + "', " + (Environment.TickCount - s_LastTimeRecorded) + ");" + Environment.NewLine);
            s_LastTimeRecorded = Environment.TickCount;
        }

        #endregion


        // 初始化脚本
        private void InitJS()
        {
            try {
                s_MainWebBrowser = new WebBrowser();
                s_MainWebBrowser.ObjectForScripting = this;
                s_MainWebBrowser.ScriptErrorsSuppressed = true;
                #region JS脚本内容区
                s_MainWebBrowser.DocumentText =
                    @"<html>
                                        <head>
                                            <script type='text/javascript'>
var isAborted = false;

/* 唤醒 */
function UnAbort() {isAborted = false; window.external.UnAbort();}
/* 挂起 */
function Abort() {isAborted = true; release(); window.external.Abort();}
/* 检查是否唤醒 */
function CheckAbort() {if(isAborted == true) { window.external.Abort(); throw new Error('Aborted');} }
/* 转换String为XML */
function stringtoXML(data){ if (window.ActiveXObject){ var doc = new ActiveXObject('Microsoft.XMLDOM'); doc.async='false'; doc.loadXML(data); } else { var parser = new DOMParser(); var doc = parser.parseFromString(data,'text/xml'); }	return doc; }
/* 释放 */
function release() { CheckAbort(); window.external.ReleaseMR();  }
/* 获取指定路径子Node数量 */
function countNodes(xpath) { CheckAbort(); return window.external.countNodes(xpath); } 
/* 开启新选项卡 */
function tabnew() { CheckAbort(); window.external.tabnew();}
/* 关闭当前选项卡  */
function tabclose() { CheckAbort(); window.external.tabclose();}
/* 关闭全部选项卡  */
function tabcloseall() { CheckAbort(); window.external.tabcloseall();}
/* 通过Url或XPath打开指定网页  */
function go(a) { CheckAbort(); window.external.go(a);}
/* 上一步 */
function back() { CheckAbort(); window.external.Back(); }
/* 下一步 */
function next() { CheckAbort(); window.external.Next(); }
/* 刷新页面 */
function reload() { CheckAbort(); window.external.Reload(); }
/* 停止页面 */
function stop() { CheckAbort(); window.external.Stop(); }
/* 休眠，若b=true则休眠到页面加载完毕，b=false则休眠a毫秒 */
function sleep(a, b) { CheckAbort(); window.external.sleep(a,b);}
/* 等待页面加载完成 */
function waitBrowser(){ CheckAbort(); window.external.waitBrowser();}
/* 退出App  */
function exit() { CheckAbort(); window.external.exit();}
/* 输出Log  */
function log(a) { CheckAbort(); window.external.log(a);}
/* 清除Log  */
function clearlog() { CheckAbort(); window.external.clearlog();}

/* 从XPath提取指定元素 */
function extract(xpath, type) {CheckAbort(); return window.external.extract(xpath, type);} 
/* 自动元素填充函数 */
function fill(a,b) { CheckAbort(); window.external.fill(a,b);}
/* winio填充函数 */
function winIOfill(a,b) { CheckAbort(); window.external.winIOfill(a,b);}
/* 自动元素填充函数 */
function filliframe(title, value) { CheckAbort(); window.external.filliframe(title, value); }   
/* 自动填充子iframe中的元素 */
function fillInputByIdAndIframe(iframeName, a, b){CheckAbort(); window.external.fillInputByIdAndIframe(iframeName, a, b);}
/* 自动元素填充函数 */
function filldropdown(xpath, value) { CheckAbort(); window.external.filldropdown(xpath, value); }
/* 点击指定XPath元素  */
function click(a) { CheckAbort(); window.external.click(a);}
/* 类型转换 */
function toObject(a) {CheckAbort(); var wrapper= document.createElement('div'); wrapper.innerHTML= a; return wrapper;}
/* 禁止/开启Flash支持 */
function blockFlash(isBlock) { CheckAbort(); window.external.BlockFlash(isBlock); }
/* 截屏 */
function takesnapshot(a) {CheckAbort(); window.external.TakeSnapshot(a);}
/* 图像识别 */
function imageToText(xpath, language) { CheckAbort(); return window.external.imgToText(xpath, language);}
/* 文件上传  */
function fileupload(a,b){CheckAbort(); window.external.FileUpload(a,b);}

/* 创建文件夹 */
function createfolder(a) { CheckAbort(); window.external.createfolder(a);}
/* 下载文件 a=Url b=本地文件夹 */
function download(a,b) {CheckAbort(); window.external.download(a,b);}
/* 获取指定文件夹内的文件列表 */
function getfiles(a) { CheckAbort(); return window.external.getfiles(a); }
/* 获取文件夹内的文件夹列表 */
function getfolders(a) { CheckAbort(); return window.external.getfolders(a); }
/* 读取一个文件  */
function read(a) { CheckAbort(); return window.external.read(a);}
/* 保存文件 a = 文件内容, b = 本地路径, c = 是否复写 (true: 复写, false: 尾部追加)  */
function save(a,b,c) { CheckAbort(); return window.external.save(a,b,c);}
/* 移除一个本地文件 */
function remove(a) { CheckAbort(); window.external.remove(a);}
/* 删除一个本地文件夹 */
function removefolder(a) {CheckAbort(); window.external.removefolder(a);}
/* 资源管理器打开一个文件夹 */
function explorer(a) { CheckAbort(); window.external.explorer(a); }

function logoff() { CheckAbort(); window.external.logoff();} 
function lockworkstation() {CheckAbort(); window.external.lockworkstation();} 
function forcelogoff() { CheckAbort(); window.external.forcelogoff();} 
function reboot() { CheckAbort(); window.external.reboot();} 
function shutdown() { CheckAbort(); window.external.shutdown();} 
function hibernate() { CheckAbort(); window.external.hibernate();} 
function standby() { CheckAbort(); window.external.standby();} 

/* 执行脚本函数  */
function excute(a) { CheckAbort(); window.external.excute(a);}
/* 执行一个进程 path=进程路径 */
function runcommand(path, parameters) { CheckAbort(); window.external.runcommand(path, parameters); }

/* 键盘鼠标控制组 */
function getCurrentMouseX() { CheckAbort(); return window.external.GetCurrentMouseX(); } 
function getCurrentMouseY() { CheckAbort(); return window.external.GetCurrentMouseY(); } 
function MouseDown(a,b) { CheckAbort(); window.external.Mouse_Down(a,b); }
function MouseUp(a,b) { CheckAbort(); window.external.Mouse_Up(a,b); }
function MouseClick(a,b) { CheckAbort(); window.external.Mouse_Click(a,b); }
function MouseDoubleClick(a,b) { CheckAbort(); window.external.Mouse_Double_Click(a,b); }
function MouseMove(a,b,c,d) {CheckAbort(); window.external.Mouse_Show(a,b,c,d); }
function MouseWheel(a,b) { CheckAbort(); window.external.Mouse_Wheel(a,b); }
function KeyDown(a,b) { CheckAbort(); window.external.Key_Down(a,b); }
function KeyUp(a,b) { CheckAbort(); window.external.Key_Up(a,b); }

function createtask(a,b,c,d,e,f) { CheckAbort(); window.external.createtask(a,b,c,d,e,f); }
function removetask(a) { CheckAbort(); window.external.removetask(a);}
function generatekeys() { CheckAbort(); window.external.generatekeys();}
function encrypt(a, b) { CheckAbort(); return window.external.encrypt(a, b);}
function decrypt(a, b) { CheckAbort(); return window.external.decrypt(a, b);}
function showpicture(a,b) { CheckAbort(); window.external.showimage(a,b); }
function savefilterimage(a) { CheckAbort(); window.external.savefilterimage(a); }

function writetextimage(a, b) {CheckAbort(); window.external.writetextimage(a,b); } 
function getcurrenturl() {CheckAbort(); return window.external.getCurrentUrl();}
function scrollto(a) {CheckAbort(); window.external.scrollto(a); }
function getheight() { CheckAbort(); return window.external.getheight(); }
function gettitle() { CheckAbort(); return window.external.gettitle(); } 
function getlinks(a) { CheckAbort(); return window.external.getlinks(a); } 
function getCurrentContent() { CheckAbort(); return window.external.getCurrentContent(); } 
function getCurrentPath() { CheckAbort(); return window.external.getCurrentPath(); } 
function checkelement(a) { CheckAbort(); return window.external.checkelement(a);}
function readCellExcel(a, b, c, d) { CheckAbort(); return window.external.readCellExcel(a,b,c,d);}
function writeCellExcel(a, b, c, d) { CheckAbort(); window.external.writeCellExcel(a,b,c,d); }
function replaceMsWord(a, b, c, d) { CheckAbort(); window.external.replaceMsWord(a,b,c,d); } 

function captchaborder(a,b) { CheckAbort(); window.external.CaptchaBorder(a,b); } 
function saveImageFromElement(a,b) { CheckAbort(); window.external.SaveImageFromElement(a,b);}
function getControlText(a,b,c) { CheckAbort(); return window.external.GetControlText(a,b,c); }
function setControlText(a,b,c,d) { CheckAbort(); window.external.SetControlText(a,b,c,d); }
function clickControl(a,b,c) { CheckAbort(); window.external.ClickControl(a,b,c); } 

function getTables(name, dbName) { CheckAbort(); return window.external.GetTables(name, dbName); }
function getColumns(name, dbName, table) { CheckAbort(); return window.external.GetColumns(name, dbName, table); }
function getRows(name, dbName, sql) { CheckAbort(); return window.external.GetRows(name, dbName, sql); }
function excuteQuery(name, dbName, sql) { CheckAbort(); return window.external.ExcuteQuery(name, dbName, sql); } 
function removeStopWords(text) { CheckAbort(); return window.external.RemoveStopWords(text); }
function addElement(path, node1, node2, text) { CheckAbort(); return window.external.AddElement(path, node1, node2, text); }
function checkXmlElement(path, node, text) { CheckAbort(); return window.external.CheckXmlElement(path, node, text); }
function getXmlElement(path, node) { CheckAbort(); return window.external.GetXmlElement(path, node); }
function getParentElement(path, node, text) { CheckAbort(); return window.external.GetParentElement(path, node, text); } 
function extractbyRegularExpression(pattern, groupName) { CheckAbort(); return window.external.ExtractUsingRegularExpression(pattern, groupName); }
function addToDownload(fileName, url, folder) { CheckAbort(); return window.external.AddToDownload(fileName, url, folder); }
function startDownload() { CheckAbort(); return window.external.StartDownload(); }                                              
                                            </script>
                                        </head>
                                        <body>
                                            
                                        </body>
                                    </html>";
                #endregion
                this.Controls.Add(s_MainWebBrowser);
            }
            catch (Exception e)
            {
                FKLog("创建浏览器并初始化JS脚本失败:" + e.ToString());
            }
        }

        public void OpenUrl(string strUrl)
        {
            if (s_CurrentTab == null)
            {
                try {
                    FKDebugLog("Before CreateNewTab");
                    CreateNewTab();
                    FKDebugLog("After CreateNewTab");
                }
                catch (Exception ex)
                {
                    throw new Exception("Create tab failed : " + ex.ToString());
                }
            }

            if (strUrl.StartsWith(" / "))
            {
                FKDebugLog("Before OpenWebBrowserByXPath");
                // 使用XPath
                OpenWebBrowserByXPath(strUrl);
                FKDebugLog("After OpenWebBrowserByXPath");
            }
            else
            {
                if (s_CurrentTab.Controls.Count > 0)
                {
                    s_CurrentTab.Controls.RemoveAt(0);
                }
                // 打开网页
                try
                {
                    FKDebugLog("Before OpenWebBrowser");
                    OpenWebBrowser(strUrl);
                    FKDebugLog("After OpenWebBrowser");
                }
                catch (Exception ex)
                {
                    throw new Exception("Open web browser failed : " + ex.ToString());
                }
            }
        }

        // 创建新页面并激活
        public void CreateNewTab()
        {
            // 创建新页面
            TabPage tab = new TabPage(Language.LanguageRes.EmptyTab);
            tabMain.Controls.Add(tab);
            tabMain.SelectedTab = tab;

            // 保存
            s_CurrentTab = tab;
        }

        // 关闭单一选项卡
        public void CloseCurTab()
        {
            if (tabMain.TabPages.Count <= 0)
            {
                s_CurrentTab = null;
                return;
            }

            if (tabMain.SelectedTab.Controls.Count > 0)
            {
                tabMain.SelectedTab.Controls[0].Dispose();
            }
            tabMain.SelectedTab.Dispose();

            if (tabMain.TabPages.Count > 1)
            {
                tabMain.SelectTab(tabMain.TabPages.Count - 1);
            }
            s_CurrentTab = tabMain.SelectedTab;
        }

        // 关闭全部选项卡
        public void CloseAllTabs()
        {
            while (tabMain.TabPages.Count > 0)
            {
                Application.DoEvents();
                if (tabMain.TabPages[0].Controls.Count > 0)
                {
                    tabMain.TabPages[0].Controls[0].Dispose();
                }
                tabMain.TabPages[0].Dispose();
            }

            s_CurrentTab = null;

            // 保持不空
            CreateNewTab();
        }

        // 点选了一个新的选项卡
        public void SelectNewTab()
        {
            if (tabMain.TabCount <= 0)
                return;

            GeckoWebBrowser GeckoWebTab = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (GeckoWebTab == null)
                return;

            tbxAddress.Text = GeckoWebTab.Url.ToString();
            string strTitle = GeckoWebTab.DocumentTitle;
            s_CurrentTab.Text = (strTitle.Length > 8 ? strTitle.Substring(0, 8) + "..." : strTitle);

            if (string.IsNullOrEmpty(s_CurrentTab.Text))
            {
                string strUrl = GeckoWebTab.Url.ToString();
                int nPos = strUrl.IndexOf("https://");
                if (nPos >= 0)
                    strUrl = strUrl.Remove(nPos, "https://".Length);

                nPos = strUrl.IndexOf("http://");
                if (nPos >= 0)
                    strUrl = strUrl.Remove(nPos, "http://".Length);
                s_CurrentTab.Text = (strUrl.Length > 8 ? strUrl.Substring(0, 8) + "..." : strUrl);
            }
        }

        // 刷新当前选项卡页面
        public void RefreshCurTab()
        {
            GeckoWebBrowser GeckoWebTab = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (GeckoWebTab == null)
                return;

            GeckoWebTab.Refresh();
        }

        // 打开“上一步”网页
        public void BackLinkCurTab()
        {
            GeckoWebBrowser GeckoWebTab = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (GeckoWebTab == null)
                return;

            GeckoWebTab.GoBack();
        }

        // 打开“下一步”网页
        public void NextLinkCurTab()
        {
            GeckoWebBrowser GeckoWebTab = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (GeckoWebTab == null)
                return;

            GeckoWebTab.GoForward();
        }

        // 显示/隐藏脚本编辑窗口
        public void ToggleScriptWindowVisible()
        {
            this.mainSplitContainer.Panel1.SuspendLayout();
            this.mainSplitContainer.Panel2.SuspendLayout();

            if (s_IsScriptWindowShow)
            {
                this.mainSplitContainer.Panel2Collapsed = true;
                this.mainSplitContainer.Panel2.Hide();
            }
            else
            {
                this.mainSplitContainer.Panel2Collapsed = false;
                this.mainSplitContainer.Panel2.Show();
            }

            this.mainSplitContainer.Panel1.ResumeLayout(false);
            this.mainSplitContainer.Panel1.PerformLayout();
            this.mainSplitContainer.Panel2.ResumeLayout(false);
            this.mainSplitContainer.Panel2.PerformLayout();

            // 焦点给予脚本框
            if (!s_IsScriptWindowShow)
            {
                if (tbxCode != null)
                    tbxCode.Focus();
            }

            // 修改值
            s_IsScriptWindowShow = !s_IsScriptWindowShow;
        }

        // 创建新脚本
        public void CreateNewScript()
        {
            // 清空编辑框
            tbxCode.Text = "";
            // 记录当前脚本文件名
            s_StrLastScriptFileName = "";
            // 修改下方提示窗文字
            toolStripStatus.Text = Language.LanguageRes.Tips_CreateScript;
        }

        // 打开新脚本文件
        public void OpenAScriptFile()
        {
            openFileDialog1.Filter = Language.LanguageRes.Filter_OpenAScriptFile;
            openFileDialog1.Multiselect = false;
            openFileDialog1.Title = Language.LanguageRes.WndTitle_OpenAScriptFile;

            if (openFileDialog1.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                string code = File.ReadAllText(openFileDialog1.FileName);
                tbxCode.Text = "";
                if (!string.IsNullOrEmpty(code))
                {
                    tbxCode.Text = code;
                }

                s_StrLastScriptFileName = openFileDialog1.FileName;
            }
            catch (Exception ex)
            {
                FKLog(string.Format(Language.LanguageRes.Log_OpenFileError, ex.ToString()));
            }
        }

        // 保存脚本文件
        public void SaveScriptFile()
        {
            toolStripStatus.Text = Language.LanguageRes.Tips_SaveScript;

            if (string.IsNullOrEmpty(s_StrLastScriptFileName))
            {
                // 另存为
                saveFileDialog1.Filter = Language.LanguageRes.Filter_SaveAScriptFile1;
                saveFileDialog1.Title = Language.LanguageRes.WndTitle_SaveAScriptFile1;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(saveFileDialog1.FileName, false))
                        {
                            file.Write(tbxCode.Text);
                        }

                        toolStripStatus.Text = Language.LanguageRes.Tips_FileSaved;
                        s_StrLastScriptFileName = saveFileDialog1.FileName;
                    }
                    catch (Exception ex)
                    {
                        FKLog(string.Format(Language.LanguageRes.Log_SaveFileError, ex.ToString()));
                    }
                }
                else
                {
                    toolStripStatus.Text = "...";
                }
            }
            else
            {
                // 保存已记录文件名
                FileInfo fileInfo = new FileInfo(s_StrLastScriptFileName);
                if (fileInfo.IsReadOnly)
                {
                    FKLog(Language.LanguageRes.Tips_FileReadOnly);
                }
                else
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(s_StrLastScriptFileName, false))
                    {
                        file.Write(tbxCode.Text);
                    }
                    toolStripStatus.Text = Language.LanguageRes.Tips_FileSaved;
                }
            }
        }

        // 脚本文件另存为
        public void SaveAsScriptFile()
        {
            s_StrLastScriptFileName = "";
            SaveScriptFile();
        }

        // 执行脚本程序
        private void ExcuteJSCode(string strCode)
        {
            s_MainWebBrowser.Document.InvokeScript("UnAbort");
            object obj = s_MainWebBrowser.Document.InvokeScript("eval", new object[] { strCode });
        }

        // 执行脚本程序
        public void RunScript()
        {
            // 正在执行，禁止二次执行
            if (!toolStripRunScript.Text.Equals(Language.LanguageRes.UI_RunScriptBtn))
                return;

            try {
                // 执行前
                s_IsScriptStop = false;
                toolStripRecordScript.Enabled = false;           // 运行状态进行进行录制
                s_MainWebBrowser.Focus();
                //s_MainWebBrowser.Document.Body.Focus();          // 强制页面获得焦点

                toolStripStatus.Text = Language.LanguageRes.Tip_BeginRunScript;
                toolStripRunScript.Text = Language.LanguageRes.UI_StopScriptBtn;
                s_MainWebBrowser.Document.InvokeScript("UnAbort");
                // 执行中
                if (!string.IsNullOrEmpty(tbxCode.Text))
                {
                    ExcuteJSCode(tbxCode.Text);
                }
                // 执行完毕，恢复
                toolStripRunScript.Text = Language.LanguageRes.UI_RunScriptBtn;
                toolStripStatus.Text = Language.LanguageRes.Tip_CompleteRunScript;
                FKLog(Language.LanguageRes.Tip_CompleteRunScript);
                toolStripRecordScript.Enabled = true;   // 允许录制
            }
            catch (Exception e)
            {
                FKLog("执行脚本失败: error = " + e.ToString());
            }
        }

        // 停止执行脚本
        public void StopScript()
        {
            s_IsScriptStop = true;
            toolStripRunScript.Text = Language.LanguageRes.UI_RunScriptBtn;
            toolStripStatus.Text = "";

            s_MainWebBrowser.Document.InvokeScript("Abort");
        }

        // 记录脚本
        public void RecordScript()
        {
            //  记录上次Hook时的时间
            s_LastTimeRecorded = Environment.TickCount;

            if (toolStripRecordScript.Text.Equals(Language.LanguageRes.UI_RecordBtn))
            {
                tbxCode.Text = "";
                toolStripRunScript.Enabled = false; // 录制时禁止运行
                s_MouseHook.Start();
                s_KeyboardHook.Start();
                toolStripRecordScript.Text = Language.LanguageRes.UI_StopScriptBtn;
            }
            else
            {
                toolStripRunScript.Enabled = true;  // 录制完毕允许运行
                s_MouseHook.Stop();
                s_KeyboardHook.Stop();
                toolStripRecordScript.Text = Language.LanguageRes.UI_RecordBtn;
            }
        }

        // 最小化窗口
        public void MinimizeWindow()
        {
            notifyIcon.BalloonTipText = Language.LanguageRes.Tips_MinAppNotify;
            notifyIcon.Visible = true;
            notifyIcon.ShowBalloonTip(200);

            this.WindowState = FormWindowState.Minimized;
            //this.Hide();
        }

        // 最小化->恢复显示窗口
        public void ShowWindow()
        {
            this.Show();
            this.WindowState = FormWindowState.Maximized;
        }

        // 打开指定Url
        void OpenWebBrowser(string strUrl)
        {
            if (String.IsNullOrEmpty(strUrl))
                return;
            if (strUrl.Equals("about:blank"))
                return;

            GeckoWebBrowser GeckoBrowser = null;
            try {
                FKDebugLog("Before InitializeLifetimeService");
                GeckoBrowser = (GeckoWebBrowser)GetCurrentWebBrowserTab();
                if (GeckoBrowser == null)
                    GeckoBrowser = new GeckoWebBrowser();
                GeckoBrowser.InitializeLifetimeService();
                FKDebugLog("After InitializeLifetimeService");

                // 重置回调函数
                GeckoBrowser.ProgressChanged -= Func_Browser_ProgressChanged;
                GeckoBrowser.ProgressChanged += Func_Browser_ProgressChanged;
                GeckoBrowser.Navigated -= Func_Browser_Navigated;
                GeckoBrowser.Navigated += Func_Browser_Navigated;
                GeckoBrowser.DocumentCompleted -= Func_Browser_DocumentCompleted;
                GeckoBrowser.DocumentCompleted += Func_Browser_DocumentCompleted;
                GeckoBrowser.CanGoBackChanged -= Func_Browser_CanGoBackChanged;
                GeckoBrowser.CanGoBackChanged += Func_Browser_CanGoBackChanged;
                GeckoBrowser.CanGoForwardChanged -= Func_Browser_CanGoForwardChanged;
                GeckoBrowser.CanGoForwardChanged += Func_Browser_CanGoForwardChanged;
                GeckoBrowser.NSSError -= Func_Browser_NSError;
                GeckoBrowser.NSSError += Func_Browser_NSError;
                GeckoBrowser.CreateWindow -= Func_Browser_CreateWindow;
                GeckoBrowser.CreateWindow += Func_Browser_CreateWindow;
                GeckoBrowser.CreateWindow2 -= Func_Browser_CreateWindow2;
                GeckoBrowser.CreateWindow2 += Func_Browser_CreateWindow2;
                // GeckoBrowser.DocumentTitleChanged  -= Func_Browser_DocumentTitleChanged;
                // GeckoBrowser.DocumentTitleChanged  += Func_Browser_DocumentTitleChanged;
            }
            catch (Exception ex)
            {
                FKLog("Init gecko browser failed : " + ex.ToString());
            }

            // 添加新Web页面
            tabMain.SelectedTab.Controls.Add(GeckoBrowser);
            GeckoBrowser.Dock = DockStyle.Fill;
            try {
                FKDebugLog("Before Navigate");
                GeckoBrowser.Navigate(strUrl);
                FKDebugLog("After Navigate");
            }
            catch (Exception ex)
            {
                FKLog("GeckoBrowser navigate : " + ex.ToString());
            }

            // 保存记录
            s_CurrentTab = tabMain.SelectedTab;
        }

        // 错误处理
        private void Func_Browser_NSError(object sender, Gecko.Events.GeckoNSSErrorEventArgs e)
        {
            // 忽略SSL安全检查
            InoredBadCert(e.Uri);

            GeckoWebBrowser GeckoBrowser = (GeckoWebBrowser)sender;
            try
            {
                // 忽略""Cannot communicate securely with peer: no common encryption algorithm(s).""错误
                if (e.Message.Contains("algorithm(s)."))
                {
                    e.Handled = true;
                    return;
                }
                GeckoBrowser.Navigate(e.Uri.AbsoluteUri);
            }
            catch (Exception ex)
            {
                FKLog("GeckoBrowser navigate : " + ex.ToString());
            }
            e.Handled = true;
        }

        // 弹出窗口
        private void Func_Browser_CreateWindow2(object sender, GeckoCreateWindow2EventArgs e)
        {
            FKDebugLog("Before Func_Browser_CreateWindow2");
            e.Cancel = true;
            if (e.Uri == "chrome://global/content/commonDialog.xul")
            {
                return;
            }

            //e.Cancel = true; // 直接忽略掉

            // 本页面打开
            //e.WebBrowser.Navigate(e.Uri);

            // 新选项卡打开  
            CreateNewTab();
            OpenUrl(e.Uri);


            // GeckoWebBrowser wb = new GeckoWebBrowser();
            //wb.Dock = DockStyle.Fill;
            //wb.CreateControl();
            //TabPage tab1 = new TabPage("New Tab");
            //tabMain.Controls.Add(tab1);
            //wb.Navigate(e.Uri);
            //wb.DocumentCompleted += Func_Browser_DocumentCompleted;
            FKDebugLog("After Func_Browser_CreateWindow2");
        }

        private void Func_Browser_CreateWindow(object sender, GeckoCreateWindowEventArgs e)
        {
            FKDebugLog("Before Func_Browser_CreateWindow");
            
            FKDebugLog("After Func_Browser_CreateWindow");
        }

        private bool InoredBadCert(Uri u)
        {
            try {
                return Gecko.CertOverrideService.RememberRecentBadCert(u);
            }
            catch
            {
                //FKLog("Inored bad cert: " + e.ToString());
                return true;
            }
            /*
            if (u == null)
                return false;

            ComPtr<nsISSLStatus> aSSLStatus = null;
            try
            {
                int nFlags = 0;
                nFlags |= nsICertOverrideServiceConsts.ERROR_MISMATCH;
                nFlags |= nsICertOverrideServiceConsts.ERROR_TIME;
                nFlags |= nsICertOverrideServiceConsts.ERROR_UNTRUSTED;
               
                int nPort = u.Port;
                if (nPort == -1) nPort = 443;
                using (var hostWithPort = new nsAString(u.Host + ":" + nPort.ToString()))
                {
                    using (var certDBSvc = Xpcom.GetService2<nsIX509CertDB>(Contracts.X509CertDb))
                    {
                        if (certDBSvc != null)
                        {
                            //using (var recentBadCerts = (certDBSvc.Instance.GetRecentBadCerts(false)).AsComPtr())
                            //{
                            //    if (recentBadCerts != null)
                            if( recentCertsSvc == null )
                            {
                                recentCertsSvc = Xpcom.GetService<nsIRecentBadCertsService>
                            }
                                    aSSLStatus = recentBadCerts.Instance.GetRecentBadCert(hostWithPort);
                           // }
                        }
                    }
                }
                using (var cert = aSSLStatus.Instance.GetServerCertAttribute().AsComPtr())
                {
                    if (aSSLStatus == null || Gecko.CertOverrideService.HasMatchingOverride(u, cert))
                    {
                        return false;
                    }
                    Gecko.CertOverrideService.RememberValidityOverride(u, cert, nFlags);
                }
            }
            finally
            {
                if (aSSLStatus != null)
                {
                    aSSLStatus.Dispose();
                }
            }
            return true;
            */
        }

        // 打开XPath路径
        void OpenWebBrowserByXPath(string xpath)
        {
            GeckoWebBrowser GeckoWB = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (GeckoWB == null)
                return;

            // 获取XPath路径元素
            GeckoHtmlElement GeckoElm = GetElementByXpath(GeckoWB.Document, xpath);
            if (GeckoElm == null)
                return;

            UpdateUrlAbsolute(GeckoWB.Document, GeckoElm);
            string strUrl = GeckoElm.GetAttribute("href");
            if (string.IsNullOrEmpty(strUrl))
                return;

            // 实际打开
            GeckoWB.Navigate(strUrl);
        }

        // 获取当前的WebBrowserTab
        private object GetCurrentWebBrowserTab()
        {
            if (tabMain.SelectedTab == null)
                return null;
            if (tabMain.SelectedTab.Controls.Count <= 0)
                return null;
            Control ctr = tabMain.SelectedTab.Controls[0];
            if (ctr == null)
            {
                return null;
            }
            return ctr as object;
        }

        // 解析XPath，获取元素
        private GeckoHtmlElement GetElementByXpath(GeckoDocument doc, string xpath)
        {
            if (doc == null)
                return null;

            xpath = xpath.Replace("/html/", "");
            GeckoElementCollection eleColec = doc.GetElementsByTagName("html");
            if (eleColec.Length == 0)
                return null;

            GeckoHtmlElement htmlElement = eleColec[0];
            string[] tagList = xpath.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string tag in tagList)
            {
                System.Text.RegularExpressions.Match mat = System.Text.RegularExpressions.Regex.Match(tag, "(?<tag>.+)\\[@id='(?<id>.+)'\\]");
                if (mat.Success == true)
                {
                    string id = mat.Groups["id"].Value;
                    GeckoHtmlElement tmpEle = doc.GetHtmlElementById(id);
                    if (tmpEle != null) htmlElement = tmpEle;
                    else
                    {
                        htmlElement = null;
                        break;
                    }
                }
                else
                {
                    mat = System.Text.RegularExpressions.Regex.Match(tag, "(?<tag>.+)\\[(?<ind>[0-9]+)\\]");
                    if (mat.Success == false)
                    {
                        GeckoHtmlElement tmpEle = null;
                        foreach (GeckoNode it in htmlElement.ChildNodes)
                        {
                            if (it.NodeName.ToLower() == tag)
                            {
                                tmpEle = (GeckoHtmlElement)it;
                                break;
                            }
                        }
                        if (tmpEle != null) htmlElement = tmpEle;
                        else
                        {
                            htmlElement = null;
                            break;
                        }
                    }
                    else
                    {
                        string tagName = mat.Groups["tag"].Value;
                        int ind = int.Parse(mat.Groups["ind"].Value);
                        int count = 0;
                        GeckoHtmlElement tmpEle = null;
                        foreach (GeckoNode it in htmlElement.ChildNodes)
                        {
                            if (it.NodeName.ToLower() == tagName)
                            {
                                count++;
                                if (ind == count)
                                {
                                    tmpEle = (GeckoHtmlElement)it;
                                    break;
                                }
                            }
                        }
                        if (tmpEle != null) htmlElement = tmpEle;
                        else
                        {
                            htmlElement = null;
                            break;
                        }
                    }
                }
            }
            return htmlElement;
        }


        protected void UpdateUrlAbsolute(GeckoDocument doc, GeckoHtmlElement ele)
        {
            string link = doc.Url.GetLeftPart(UriPartial.Authority);

            var eleColec = ele.GetElementsByTagName("IMG");
            foreach (GeckoHtmlElement it in eleColec)
            {
                if (!it.GetAttribute("src").StartsWith("http://"))
                {
                    it.SetAttribute("src", link + it.GetAttribute("src"));
                }
            }
            eleColec = ele.GetElementsByTagName("A");
            foreach (GeckoHtmlElement it in eleColec)
            {
                if (!it.GetAttribute("href").StartsWith("http://"))
                {
                    it.SetAttribute("href", link + it.GetAttribute("href"));
                }
            }
        }

        // 进度条更新函数
        void Func_Browser_ProgressChanged(object sender, GeckoProgressEventArgs e)
        {
            try {
                progressbar.Maximum = (int)e.MaximumProgress;
                var currentProgress = (int)e.CurrentProgress;
                if (currentProgress <= progressbar.Maximum)
                {
                    progressbar.Value = (int)e.CurrentProgress;
                }
            }
            catch { }
        }

        // 更新实际地址
        void Func_Browser_Navigated(object sender, GeckoNavigatedEventArgs e)
        {
            FKDebugLog("Before Func_Browser_Navigated");
            string url = string.Empty;
            url = ((GeckoWebBrowser)sender).Url.ToString();
            if (url != "about:blank")
            {
                tbxAddress.Text = url;
            }
            FKDebugLog("After Func_Browser_Navigated");
        }

        // 加载页面完成回调
        void Func_Browser_DocumentCompleted(object sender, EventArgs e)
        {
            FKDebugLog("Before Func_Browser_DocumentCompleted");
            GeckoWebBrowser GeckoBrowser = (GeckoWebBrowser)sender;

            string strTitle = GeckoBrowser.DocumentTitle;
            s_CurrentTab.Text = (strTitle.Length > 12 ? strTitle.Substring(0, 12) + "..." : strTitle);
            if (string.IsNullOrEmpty(s_CurrentTab.Text))
            {
                string strUrl = GeckoBrowser.Url.ToString();
                int nPos = strUrl.IndexOf("https://");
                if (nPos >= 0)
                    strUrl = strUrl.Remove(nPos, "https://".Length);

                nPos = strUrl.IndexOf("http://");
                if (nPos >= 0)
                    strUrl = strUrl.Remove(nPos, "http://".Length);
                s_CurrentTab.Text = (strUrl.Length > 12 ? strUrl.Substring(0, 12) + "..." : strUrl);
            }
            if (!GeckoBrowser.Url.ToString().Equals("about:blank"))
                tbxAddress.Text = GeckoBrowser.Url.ToString();

            GeckoBrowser.DomContextMenu += Func_Browser_DomContextMenu;
            GeckoBrowser.NoDefaultContextMenu = true;

            // 删除当前连接
            GeckoBrowser.DocumentCompleted -= Func_Browser_DocumentCompleted;

            // 通知页面加载完成
            if (!GeckoBrowser.IsBusy)
            {
                s_IsPageReady = true;
            }
            toolStripStatus.Text = "页面打开完成...";
            FKDebugLog("After Func_Browser_DocumentCompleted");

            // 输出页面源代码
            // FKDebugLog(GeckoBrowser.Document.GetElementsByTagName("html").ElementAt(0).InnerHtml);
        }

        // Dom页面鼠标按键消息
        void Func_Browser_DomContextMenu(object sender, DomMouseEventArgs e)
        {
            // 如果按下右键
            if (e.Button.ToString().IndexOf("Right") != -1)
            {
                // 开启右键菜单显示
                s_tagRightClickPoint = Cursor.Position; // 保存当前鼠标位置
                contextMenuBrowser.Show(Cursor.Position);

                GeckoWebBrowser GeckoBrowser = (GeckoWebBrowser)GetCurrentWebBrowserTab();
                if (GeckoBrowser != null)
                {
                    // 记录当前鼠标右键选择的元素对象
                    s_HtmlElement = GeckoBrowser.Document.ElementFromPoint(e.ClientX, e.ClientY);
                }
            }
        }

        // 后退键是否有效
        void Func_Browser_CanGoBackChanged(object sender, EventArgs e)
        {
            var GeckoBrowser = (GeckoWebBrowser)sender;
            if (GeckoBrowser != null)
            {
                toolStripBack.Enabled = GeckoBrowser.CanGoBack;
            }
            else
            {
                toolStripBack.Enabled = false;
            }
        }

        // 前进键是否有效
        void Func_Browser_CanGoForwardChanged(object sender, EventArgs e)
        {
            var GeckoBrowser = (GeckoWebBrowser)sender;
            if (GeckoBrowser != null)
            {
                toolStripNext.Enabled = GeckoBrowser.CanGoForward;
            }
            else
            {
                toolStripNext.Enabled = false;
            }
        }

        // 窗口标题发生更变消息
        /*
        void Func_Browser_DocumentTitleChanged(object sender, EventArgs e)
        {
            var GeckoBrowser = (GeckoWebBrowser)sender;
            if (GeckoBrowser != null)
            {
                string strTitle = GeckoBrowser.DocumentTitle;
                s_CurrentTab.Text = (strTitle.Length > 10 ? strTitle.Substring(0, 10) + "..." : strTitle);
            }
        }
        */

        // 在地址栏输入按键事件
        private void tbxAddress_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue != 13)   // 只处理回车
                return;

            this.btnGo_Click(sender, e);
        }

        // 创建新选项卡按钮点击
        private void btnNewTab_Click(object sender, EventArgs e)
        {
            CreateNewTab();
        }

        // 按下“访问GO”按钮
        private void btnGo_Click(object sender, EventArgs e)
        {
            try {
                FKDebugLog("Before OpenUrl - click");
                OpenUrl(tbxAddress.Text);
                FKDebugLog("After OpenUrl - click");
            }
            catch (Exception ex)
            {
                FKLog("Open url failed: " + ex.ToString());
            }
        }

        // 关闭单一选项卡按钮
        private void btnCloseTab_Click(object sender, EventArgs e)
        {
            CloseCurTab();
        }

        // 关闭全部选项卡按钮
        private void btnCloseAllTab_Click(object sender, EventArgs e)
        {
            CloseAllTabs();
        }

        // 点选了另外一个选项卡
        private void tabMain_Selected(object sender, TabControlEventArgs e)
        {
            SelectNewTab();
        }

        // 点选刷新按钮
        private void toolStripReload_Click(object sender, EventArgs e)
        {
            RefreshCurTab();
        }

        // 网页后退按钮按下
        private void toolStripBack_Click(object sender, EventArgs e)
        {
            BackLinkCurTab();
        }

        // 网页前进按钮按下
        private void toolStripNext_Click(object sender, EventArgs e)
        {
            NextLinkCurTab();
        }

        // 点选隐藏/显示脚本编辑栏
        private void scriptWindowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToggleScriptWindowVisible();
        }

        // 创建新脚本
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateNewScript();
        }

        // 打开新脚本文件
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenAScriptFile();
        }

        // 保存脚本按钮按下
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveScriptFile();
        }

        // 另存为按钮按下
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAsScriptFile();
        }

        // 最小化窗口按键按下
        private void hideBrowserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MinimizeWindow();
        }

        // 窗口最小ICON被点击
        private void notifyIcon_Click(object sender, EventArgs e)
        {
            ShowWindow();
        }

        // 窗口初始化函数
        private void Form_Main_Load(object sender, EventArgs e)
        {
            AppInit();
            OnWindowInit();
        }

        // 简体中文按钮
        private void zhCNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateLanguage("zh-CN");
        }

        // 英文按钮
        private void enUSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateLanguage("en-US");
        }

        // 按下开启关闭Flash按钮
        private void menuItemDisableFlash_Click(object sender, EventArgs e)
        {
            // TODO：
        }

        // 新建脚本按钮
        private void toolStripNewScript_Click(object sender, EventArgs e)
        {
            CreateNewScript();
        }

        // 打开脚本按钮
        private void toolStripOpenScript_Click(object sender, EventArgs e)
        {
            OpenAScriptFile();
        }

        // 保存脚本按钮
        private void toolStripSaveScript_Click(object sender, EventArgs e)
        {
            SaveScriptFile();
        }

        // 脚本另存为按钮
        private void toolStripSaveScriptAs_Click(object sender, EventArgs e)
        {
            SaveAsScriptFile();
        }

        // 清空脚本编辑器文本按钮
        private void toolStripClearScript_Click(object sender, EventArgs e)
        {
            tbxCode.Text = "";
        }

        // 运行脚本按钮按下
        private void toolStripRunScript_Click(object sender, EventArgs e)
        {
            if (toolStripRunScript.Text.Equals(Language.LanguageRes.UI_RunScriptBtn))
                RunScript();
            else
                StopScript();
        }

        // 记录脚本按钮按下
        private void toolStripRecordScript_Click(object sender, EventArgs e)
        {
            RecordScript();
        }

        // 关闭窗口
        private void Form_Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            Application.Exit();
            Process.GetCurrentProcess().Kill();
        }

        // “关于” 按钮按下
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_About dlg = new Form_About();
            dlg.ShowDialog(this);
        }

        // JS函数之 sleep
        public void sleep(int seconds, bool isBreakWhenWBCompleted)
        {
            for (int i = 0; i < seconds; i++)
            {
                if (s_IsScriptStop == false)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(100);

                    toolStripStatus.Text = "Sleep: " + ((i + 1) * 100) + "/" + (seconds * 100);
                }
                else
                {
                    break;
                }
            }
            toolStripStatus.Text = "";
        }

        // JS函数之 go
        public void go(string url)
        {
            toolStripStatus.Text = "正在打开页面，请稍等...";
            s_IsPageReady = false;
            tbxAddress.Text = url;  // 提前显示地址
            FKDebugLog("Before OpenUrl - go script");
            OpenUrl(url);
            FKDebugLog("Before OpenUrl - go script");
            waitBrowser();
        }

        public void waitBrowser()
        {
            GeckoWebBrowser wb = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (wb == null)
                return;
            TimeSpan begin = new TimeSpan(DateTime.Now.Ticks);
            while (!s_IsPageReady)
            {
                TimeSpan cur = new TimeSpan(DateTime.Now.Ticks);
                TimeSpan diff = cur.Subtract(begin).Duration();
                // 1分钟超时
                if (diff.Days > 0 || diff.Hours > 0 || diff.Minutes >= 1)
                {
                    FKLog("Open url time out, more than 1 minute.");
                    break;
                }
                try
                {
                    Application.DoEvents();
                }
                catch { }
            }
            s_IsPageReady = false;
        }

        // JS函数之log
        public void log(string strText)
        {
            FKLog(strText);
        }

        // JS函数之 extract 元素提取
        public string extract(string xpath, string type)
        {
            string result = string.Empty;
            GeckoHtmlElement elm = null;

            GeckoWebBrowser wb = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (wb == null)
                return Language.LanguageRes.Log_CanntFindWebBrowser;

            if (xpath.StartsWith("/"))
            {
                elm = GetElementByXpath(wb.Document, xpath);
                if (elm != null)
                    UpdateUrlAbsolute(wb.Document, elm);
            }
            else
            {
                var id = xpath;
                elm = wb.Document.GetHtmlElementById(id);
                if (elm != null)
                    UpdateUrlAbsolute(wb.Document, elm);
            }

            if (elm == null)
                return Language.LanguageRes.Log_CanntFindElm;

            switch (type)
            {
                case "html":
                    result = "HTML = " + elm.OuterHtml;
                    break;
                case "src":
                    result = "SRC = " + elm.GetAttribute("src").Trim();
                    break;
                case "text":
                    if (elm.GetType().Name == "GeckoTextAreaElement")
                    {
                        result = "TEXT = " + ((GeckoTextAreaElement)elm).Value;
                    }
                    else
                    {
                        result = "TEXT = " + elm.TextContent.Trim();
                    }
                    break;
                case "href":
                    result = "HREF = " + elm.GetAttribute("href").Trim();
                    break;
                default:
                    result = type + " = " + elm.GetAttribute(type).Trim();
                    break;
            }

            return result;
        }

        // 脚本函数-自动填充输入框
        public void fill(string id, string value)
        {
            GeckoWebBrowser wb = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (wb == null)
                return;

            if (id.StartsWith("/"))
            {
                string xpath = id;
                GeckoHtmlElement elm = GetElementByXpath(wb.Document, xpath);
                if (elm != null)
                {
                    switch (elm.TagName)
                    {
                        case "IFRAME":
                            foreach (GeckoWindow ifr in wb.Window.Frames)
                            {
                                if (ifr.Document == elm.DOMElement)
                                {
                                    ifr.Document.TextContent = value;
                                    break;
                                }
                            }
                            break;
                        case "INPUT":
                            GeckoInputElement input = (GeckoInputElement)elm;
                            //input.DefaultValue = "";
                            input.Value = value;
                            input.Focus();
                            break;
                        default:
                            break;
                    }
                }
            }
            else
            {
                Byte[] bytes = Encoding.UTF32.GetBytes(value);
                StringBuilder asAscii = new StringBuilder();
                for (int idx = 0; idx < bytes.Length; idx += 4)
                {
                    uint codepoint = BitConverter.ToUInt32(bytes, idx);
                    if (codepoint <= 127)
                        asAscii.Append(Convert.ToChar(codepoint));
                    else
                        asAscii.AppendFormat("\\u{0:x4}", codepoint);
                }

                using (AutoJSContext context = new AutoJSContext(wb.Window.JSContext))
                {
                    context.EvaluateScript("document.getElementById('" + id + "').value = '" + asAscii.ToString() + "';");
                    context.EvaluateScript("document.getElementById('" + id + "').scrollIntoView();");
                }
            }
        }

        public void winIOfill(string a, string value)
        {
            string strParam = a + "@" + value;
            string strResult;
            if (!RunCmd("FKWinIOInput.exe", strParam, out strResult))
                FKLog("WinIO result = " + strResult);
        }

        public bool RunCmd(string cmdExe, string cmdStr, out string strResult)
        {
            Process myPro = new Process();
            bool bResult = false;
            try
            {
                ProcessStartInfo info = new ProcessStartInfo(cmdExe, cmdStr);
                myPro.StartInfo.UseShellExecute = false;
                myPro.StartInfo.RedirectStandardError = true;
                myPro.StartInfo.RedirectStandardInput = true;
                myPro.StartInfo.RedirectStandardOutput = true;
                myPro.StartInfo.CreateNoWindow = true;
                myPro.StartInfo = info;

                myPro.Start();
                myPro.StandardInput.WriteLine(cmdStr);
                strResult = myPro.StandardOutput.ReadToEnd();
                myPro.StandardInput.Flush();
                myPro.WaitForExit();
                myPro.Close();

                bResult = true;
            }
            catch (Exception e)
            {
                strResult = e.Message;
            }
            myPro.Close();
            return bResult;
        }

        // 脚本函数-自动填充下拉列表
        public void filldropdown(string id, string value)
        {
            GeckoWebBrowser wb = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (wb == null)
                return;
            if (id.StartsWith("/"))
            {
                string xpath = id;
                GeckoHtmlElement elm = GetElementByXpath(wb.Document, xpath);
                if (elm == null)
                    return;
                elm.SetAttribute("selectedIndex", value);

                elm.Focus();
            }
            else
            {
                using (AutoJSContext context = new AutoJSContext(wb.Window.JSContext))
                {
                    string javascript = string.Empty;
                    context.EvaluateScript("document.getElementById('" + id + "').selectedIndex = " + value + ";");
                    JQueryExecutor jquery = new JQueryExecutor(wb.Window);
                    jquery.ExecuteJQuery("$('#" + id + "').trigger('change');");
                    context.EvaluateScript("document.getElementById('" + id + "').scrollIntoView();");
                }
            }
        }

        // 脚本函数-自动填充iFrame
        public void filliframe(string title, string value)
        {
            GeckoWebBrowser wb = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (wb == null)
                return;

            foreach (GeckoWindow ifr in wb.Window.Frames)
            {
                if (ifr.Document.Title == title)
                {
                    foreach (var item in ifr.Document.ChildNodes)
                    {
                        if (item.NodeName == "HTML")
                        {
                            foreach (var it in item.ChildNodes)
                            {
                                if (it.NodeName == "BODY")
                                {
                                    GeckoBodyElement elem = (GeckoBodyElement)it;
                                    elem.InnerHtml = value;
                                    elem.Focus();
                                }
                            }
                            break;
                        }
                    }
                    break;
                }
            }
        }
        // 脚本函数-自动填充iFrame的Input
        public void fillInputByIdAndIframe(string title, string strInputID, string value)
        {
            GeckoWebBrowser wb = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (wb == null)
                return;
            var iFrame = wb.Document.GetElementsByTagName("iframe").FirstOrDefault() as Gecko.DOM.GeckoIFrameElement;
            if (iFrame == null)
                return;

            GeckoElement element = iFrame.ContentDocument.GetElementById(strInputID);
            if (element == null)
                return;
            GeckoInputElement input = element as GeckoInputElement;
            if (input == null)
                return;
            input.Value = value;
        }

        // 脚本函数-自动点击指定元素
        public void click(string id)
        {
            GeckoWebBrowser wb = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            if (wb == null)
                return;

            if (id.StartsWith("/"))
            {
                string xpath = id;
                GeckoHtmlElement elm = GetElementByXpath(wb.Document, xpath);
                if (elm != null)
                    elm.Click();
            }
            else
            {
                using (AutoJSContext context = new AutoJSContext(wb.Window.JSContext))
                {
                    context.EvaluateScript("document.getElementById('" + id + "').click();");
                    context.EvaluateScript("document.getElementById('" + id + "').scrollIntoView();");
                }
            }
        }

        public void Mouse_Down(string mouseButton, int LastTime)
        {
            WaitApp(LastTime);

            if (mouseButton == "Left")
            {
                MouseKeyboardLibrary.MouseSimulator.MouseDown(MouseKeyboardLibrary.MouseButton.Left);
            }
            else if (mouseButton == "Right")
            {
                MouseKeyboardLibrary.MouseSimulator.MouseDown(MouseKeyboardLibrary.MouseButton.Right);
            }
            else if (mouseButton == "Middle")
            {
                MouseKeyboardLibrary.MouseSimulator.MouseDown(MouseKeyboardLibrary.MouseButton.Middle);
            }
        }

        public void Mouse_Up(string mouseButton, int LastTime)
        {
            WaitApp(LastTime);

            if (mouseButton == "Left")
            {
                MouseKeyboardLibrary.MouseSimulator.MouseUp(MouseKeyboardLibrary.MouseButton.Left);
            }
            else if (mouseButton == "Right")
            {
                MouseKeyboardLibrary.MouseSimulator.MouseUp(MouseKeyboardLibrary.MouseButton.Right);
            }
            else if (mouseButton == "Middle")
            {
                MouseKeyboardLibrary.MouseSimulator.MouseUp(MouseKeyboardLibrary.MouseButton.Middle);
            }

        }

        public void Mouse_Click(string mouseButton, int LastTime)
        {
            WaitApp(LastTime);

            if (mouseButton == "Left")
            {
                MouseKeyboardLibrary.MouseSimulator.Click(MouseKeyboardLibrary.MouseButton.Left);
            }
            else if (mouseButton == "Right")
            {
                MouseKeyboardLibrary.MouseSimulator.Click(MouseKeyboardLibrary.MouseButton.Right);
            }
            else if (mouseButton == "Middle")
            {
                MouseKeyboardLibrary.MouseSimulator.Click(MouseKeyboardLibrary.MouseButton.Middle);
            }
        }

        public void Mouse_Double_Click(string mouseButton, int LastTime)
        {
            WaitApp(LastTime);

            if (mouseButton == "Left")
            {
                MouseKeyboardLibrary.MouseSimulator.DoubleClick(MouseKeyboardLibrary.MouseButton.Left);
            }
            else if (mouseButton == "Right")
            {
                MouseKeyboardLibrary.MouseSimulator.DoubleClick(MouseKeyboardLibrary.MouseButton.Right);
            }
            else if (mouseButton == "Middle")
            {
                MouseKeyboardLibrary.MouseSimulator.DoubleClick(MouseKeyboardLibrary.MouseButton.Middle);
            }

        }

        public void Mouse_Show(int x, int y, bool isShow, int LastTime)
        {
            WaitApp(LastTime);

            MouseKeyboardLibrary.MouseSimulator.X = x;
            MouseKeyboardLibrary.MouseSimulator.Y = y;

            if (isShow)
            {
                MouseKeyboardLibrary.MouseSimulator.Show();
            }
            else if (isShow == false)
            {
                MouseKeyboardLibrary.MouseSimulator.Hide();
            }
        }

        public void Mouse_Wheel(int delta, int LastTime)
        {
            WaitApp(LastTime);

            MouseKeyboardLibrary.MouseSimulator.MouseWheel(delta);
        }

        public void Key_Down(string key, int LastTime)
        {
            WaitApp(LastTime);

            KeysConverter k = new KeysConverter();
            Keys mykey = (Keys)k.ConvertFromString(key);
            MouseKeyboardLibrary.KeyboardSimulator.KeyDown(mykey);
        }

        public void Key_Up(string key, int LastTime)
        {
            WaitApp(LastTime);

            KeysConverter k = new KeysConverter();
            Keys mykey = (Keys)k.ConvertFromString(key);
            MouseKeyboardLibrary.KeyboardSimulator.KeyUp(mykey);
        }

        // 脚本函数-APP休眠挂起
        private void WaitApp(int seconds)
        {
            Application.DoEvents();
            System.Threading.Thread.Sleep(seconds);
        }

        // 脚本函数-暂停脚本
        public void Abort()
        {
            s_IsScriptStop = true;
        }

        // 脚本函数-解除暂停
        public void UnAbort()
        {
            s_IsScriptStop = false;
        }

        // 脚本函数-创建文件夹
        public void createfolder(string path)
        {
            if (System.IO.Directory.Exists(path) == false)
                System.IO.Directory.CreateDirectory(path);
        }

        // 脚本函数-获取文件列表
        public string getfiles(string path)
        {
            string result = "";
            string r = "";
            string[] filePaths = Directory.GetFiles(path);
            foreach (string f in filePaths)
            {
                r += f + ",";
            }
            if (!string.IsNullOrEmpty(r))
            {
                result = r.Substring(0, r.Length - 1);
            }
            else
            {
                result = r;
            }
            return result;
        }

        // 脚本函数-获取文件夹列表
        public string getfolders(string path)
        {
            string result = "";
            string r = "";
            string[] directoryPaths = Directory.GetDirectories(path);
            foreach (string f in directoryPaths)
            {
                r += f + ",";
            }
            if (!string.IsNullOrEmpty(r))
            {
                result = r.Substring(0, r.Length - 1);
            }
            else
            {
                result = r;
            }
            return result;
        }

        // 脚本函数-删除单一文件
        public void remove(string path)
        {
            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
            }
        }

        // 脚本函数-删除一个文件夹

        public void removefolder(string path)
        {
            if (System.IO.Directory.Exists(path))
            {
                System.IO.Directory.Delete(path);
            }
        }

        // 脚本函数-保存一个文件
        public void save(string content, string path, bool isOverride)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path, isOverride))
            {
                file.WriteLine(content);
            }
        }

        // 脚本函数-在资源管理器中打开一个文件
        public void explorer(string path)
        {
            string argument = "/select, \"" + path + "\"";
            System.Diagnostics.Process.Start("explorer.exe", argument);
        }

        // 脚本函数-获取当前路径
        public string getCurrentPath()
        {
            return GetCurrentPath();
        }

        // 脚本函数-读取一个文件
        public string read(string path)
        {
            return Read(path);
        }


        // 脚本函数-执行命令行
        public void runcommand(string path, string parameters)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WorkingDirectory = getCurrentPath();
            startInfo.FileName = path;
            startInfo.Arguments = parameters;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            Process p = Process.Start(startInfo);
            p.WaitForExit();
        }

        // 脚本函数-图像识别
        public string imgToText(string xpath, string language)
        {
            string data = string.Empty;
            string path = string.Empty;
            path = Application.StartupPath + "\\image.jpg";
            bool isSaveSuccess = SaveImage(xpath, path);

            if (isSaveSuccess)
            {
                // DELETE LOGIC CODE
            }
            return data;
        }

        // 脚本函数-屏幕截屏
        public void TakeSnapshot(string location)
        {
            GeckoWebBrowser wbBrowser = (GeckoWebBrowser)GetCurrentWebBrowserTab();
            ImageCreator creator = new ImageCreator(wbBrowser);
            byte[] rs = creator.CanvasGetPngImage((uint)wbBrowser.Document.ActiveElement.ScrollWidth,
                (uint)wbBrowser.Document.ActiveElement.ScrollHeight);
            MemoryStream ms = new MemoryStream(rs);
            Image returnImage = Image.FromStream(ms);

            returnImage.Save(location);
        }

        // 脚本函数-加载Excel
        public string readCellExcel(string filePath, string isheetname, int irow, int icolumn)
        {
            string result = "";
            try
            {
                using (StreamReader input = new StreamReader(filePath))
                {
                    NPOI.HSSF.UserModel.HSSFWorkbook workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(new NPOI.POIFS.FileSystem.POIFSFileSystem(input.BaseStream));
                    if (null == workbook)
                    {
                        result = "";
                    }
                    NPOI.HSSF.UserModel.HSSFFormulaEvaluator formulaEvaluator = new NPOI.HSSF.UserModel.HSSFFormulaEvaluator(workbook);
                    NPOI.HSSF.UserModel.HSSFDataFormatter dataFormatter = new NPOI.HSSF.UserModel.HSSFDataFormatter(new CultureInfo("zh-CN"));

                    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheet(isheetname);
                    NPOI.SS.UserModel.IRow row = sheet.GetRow(irow);

                    if (row != null)
                    {
                        short minColIndex = row.FirstCellNum;
                        short maxColIndex = row.LastCellNum;

                        if (icolumn >= minColIndex || icolumn <= maxColIndex)
                        {
                            NPOI.SS.UserModel.ICell cell = row.GetCell(icolumn);
                            if (cell != null)
                            {
                                result = dataFormatter.FormatCellValue(cell, formulaEvaluator);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string test = ex.ToString();
            }

            return result;
        }

        // 脚本函数-写入Excel
        public void writeCellExcel(string filePath, string sheetname, string cellName, string value)
        {
            return;
        }

        private void testToolStripButton_Click(object sender, EventArgs e)
        {
            Random rm = new Random();
            SetHttpProxy(ipToolStripTextBox.Text, portToolStripTextBox.Text);
        }

        // 日志文本更变行为
        private void tbxLog_TextChanged(object sender, EventArgs e)
        {
            tbxLog.SelectionStart = tbxLog.Text.Length;
            tbxLog.ScrollToCaret();
        }

        // 双击TPP消息
        private void listViewClients_DoubleClick(object sender, EventArgs e)
        {
            //DELETE LOGIC CODE, RUN SCRIPT
        }

        private void listViewClientsMP_DoubleClick(object sender, EventArgs e)
        {
            //DELETE LOGIC CODE, RUN SCRIPT
        }

        public void AutoDo(/* params */)
        {
            try
            {
                //DELETE LOGIC CODE, PARSER SCRIPT
            }
            catch
            {

            }
        }

        // 搜索按钮
        private void button_UseFilter_Click(object sender, EventArgs e)
        {
            // DELETE LOGIC CODE, RESORT LIST
        }
    }

    public class SortList : IComparer
    {
        public int Compare(object x, object y)
        {
            try {
                return 0;
            }
            catch { return 1; }
        }
    }
}
