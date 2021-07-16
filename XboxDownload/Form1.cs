using NetFwTypeLib;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace XboxDownload
{
    public partial class Form1 : Form
    {
        internal static Boolean bServiceFlag = false, bRecordLog = true;
        internal static ConcurrentDictionary<String, Byte[]> dicDomain = new ConcurrentDictionary<String, Byte[]>();
        internal static List<Market> lsMarket = new List<Market>();
        internal static float dpixRatio = 1;
        private readonly DataTable dtDomain = new DataTable("Domain");
        private readonly DnsListen dnsListen;
        private readonly HttpListen httpListen;
        
        private readonly String domainPath = Application.StartupPath + "\\Domain";

        public Form1()
        {
            InitializeComponent();

            Form1.dpixRatio = Environment.OSVersion.Version.Major >= 10 ? CreateGraphics().DpiX / 96 : Program.Utility.DpiX / 96;
            if (Form1.dpixRatio > 1)
            {
                foreach (ColumnHeader col in lvLog.Columns)
                    col.Width = (int)(col.Width * Form1.dpixRatio);
                dgvIpList.RowHeadersWidth = (int)(dgvIpList.RowHeadersWidth * Form1.dpixRatio);
                foreach (DataGridViewColumn col in dgvIpList.Columns)
                    col.Width = (int)(col.Width * Form1.dpixRatio);
                dgvHosts.RowHeadersWidth = (int)(dgvHosts.RowHeadersWidth * Form1.dpixRatio);
                foreach (DataGridViewColumn col in dgvHosts.Columns)
                    col.Width = (int)(col.Width * Form1.dpixRatio);
                dgvDevice.RowHeadersWidth = (int)(dgvDevice.RowHeadersWidth * Form1.dpixRatio);
                foreach (DataGridViewColumn col in dgvDevice.Columns)
                    col.Width = (int)(col.Width * Form1.dpixRatio);
                foreach (ColumnHeader col in lvGame.Columns)
                    col.Width = (int)(col.Width * Form1.dpixRatio);
            }

            dnsListen = new DnsListen(this);
            httpListen = new HttpListen(this);

            tbDnsIP.Text = Properties.Settings.Default.DnsIP;
            tbComIP.Text = Properties.Settings.Default.ComIP;
            tbCnIP.Text = Properties.Settings.Default.CnIP;
            tbAppIP.Text = Properties.Settings.Default.AppIP;
            tbEaIP.Text = Properties.Settings.Default.EaIP;
            ckbRedirect.Checked = Properties.Settings.Default.Redirect;
            ckbTruncation.Checked = Properties.Settings.Default.Truncation;
            ckbLocalUpload.Checked = Properties.Settings.Default.LocalUpload;
            if (string.IsNullOrEmpty(Properties.Settings.Default.LocalPath))
                Properties.Settings.Default.LocalPath = Application.StartupPath + "\\Upload";
            tbLocalPath.Text = Properties.Settings.Default.LocalPath;
            cbListenIP.SelectedIndex = Properties.Settings.Default.ListenIP;
            ckbDnsService.Checked = Properties.Settings.Default.DnsService;
            ckbHttpService.Checked = Properties.Settings.Default.HttpService;
            ckbMicrosoftStore.Checked = Properties.Settings.Default.MicrosoftStore;

            IPAddress[] ipAddresses = Array.FindAll(Dns.GetHostEntry(string.Empty).AddressList, a => a.AddressFamily == AddressFamily.InterNetwork);
            cbLocalIP.Items.AddRange(ipAddresses);
            if (cbLocalIP.Items.Count >= 1)
            {
                int index = 0;
                if (!string.IsNullOrEmpty(Properties.Settings.Default.LocalIP))
                {
                    for (int i = 0; i < cbLocalIP.Items.Count; i++)
                    {
                        if (cbLocalIP.Items[i].ToString() == Properties.Settings.Default.LocalIP)
                        {
                            index = i;
                            break;
                        }
                    }
                }
                cbLocalIP.SelectedIndex = index;
            }

            dtDomain.Columns.Add("Enable", typeof(Boolean));
            dtDomain.Columns.Add("Domain", typeof(String));
            dtDomain.Columns.Add("IPv4", typeof(String));
            dtDomain.Columns.Add("Remark", typeof(String));
            if (File.Exists(domainPath))
            {
                try
                {
                    dtDomain.ReadXml(domainPath);
                }
                catch { }
                dtDomain.AcceptChanges();
            }
            dgvHosts.DataSource = dtDomain;
            AddDomain();

            Form1.lsMarket.AddRange((new List<Market>
            {
                new Market("新加坡", "SG", "zh-SG"),
                new Market("香港", "HK", "zh-HK"),
                new Market("台湾", "TW", "zh-TW"),
                new Market("日本", "JP", "ja-JP"),
                new Market("美国", "US", "en-US"),

                new Market("阿根廷", "AR", "es-AR"),
                new Market("阿联酋", "AE", "ar-AE"),   //en-AE
                new Market("爱尔兰" ,"IE", "en-IE"),
                new Market("奥地利", "AT", "de-AT"),
                new Market("澳大利亚", "AU", "en-AU"),
                new Market("巴西", "BR", "pt-BR"),
                new Market("比利时", "BE", "nl-BE"),
                new Market("波兰", "PL", "pl-PL"),
                new Market("丹麦", "DK", "da-DK"),
                new Market("德国", "DE", "de-DE"),
                new Market("俄罗斯", "RU", "ru-RU"),
                new Market("法国", "FR", "fr-FR"),
                new Market("芬兰", "FI", "fi-FI"),
                new Market("哥伦比亚", "CO", "es-CO"),
                new Market("韩国", "KR", "ko-KR"),
                new Market("荷兰", "NL", "nl-NL"),
                new Market("加拿大", "CA", "en-CA"),
                new Market("捷克共和国", "CZ", "cs-CZ"),
                //new Market("美国", "US", "en-US"),
                new Market("墨西哥", "MX", "es-MX"),
                new Market("南非", "ZA", "en-ZA"),
                new Market("挪威", "NO", "nb-NO"),
                new Market("葡萄牙", "PT", "pt-PT"),
                //new Market("日本", "JP", "ja-JP"),
                new Market("瑞典", "SE", "sv-SE"),
                new Market("瑞士", "CH", "de-CH"),    //fr-CH
                new Market("沙特阿拉伯", "SA", "en-SA"), //ar-SA
                new Market("斯洛伐克", "SK", "sk-SK"),
                //new Market("台湾", "TW", "zh-TW"),
                new Market("土尔其", "TR", "tr-TR"),
                new Market("西班牙", "ES", "es-ES"),
                new Market("希腊", "GR", "el-GR"),
                //new Market("香港", "HK", "zh-HK"),
                //new Market("新加坡", "SG", "zh-SG"),
                new Market("新西兰", "NZ", "en-NZ"),
                new Market("匈牙利", "HU", "en-HU"),   //hu-HU
                new Market("以色列", "IL", "en-IL"),   //he-IL
                new Market("意大利", "IT", "it-IT"),
                new Market("印度", "IN", "en-IN"),
                new Market("英国", "GB", "en-GB"),
                new Market("智利", "CL", "es-CL"),
                new Market("中国", "CN", "zh-CN")
            }).ToArray());
            cbGameMarket.Items.AddRange(Form1.lsMarket.ToArray());
            cbGameMarket.SelectedIndex = 0;
            pbGame.Image = pbGame.InitialImage;

            LinkRefreshDrive_LinkClicked(null, null);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ToolTip toolTip1 = new ToolTip
            {
                AutoPopDelay = 30000,
                IsBalloon = true
            };
            toolTip1.SetToolTip(this.labelCom, "包括以下com游戏下载域名\nassets1.xboxlive.com\nassets2.xboxlive.com\n7.assets1.xboxlive.com\ndlassets.xboxlive.com\ndlassets2.xboxlive.com\nd1.xboxlive.com\nd2.xboxlive.com\nxvcf1.xboxlive.com\nxvcf2.xboxlive.com");
            toolTip1.SetToolTip(this.labelCn, "包括以下cn游戏下载域名\nassets1.xboxlive.cn\nassets2.xboxlive.cn\ndlassets.xboxlive.cn\ndlassets2.xboxlive.cn\nd1.xboxlive.cn\nd2.xboxlive.cn");
            toolTip1.SetToolTip(this.labelApp, "包括以下应用下载域名\ndl.delivery.mp.microsoft.com\ntlu.dl.delivery.mp.microsoft.com");
            toolTip1.SetToolTip(this.labelEA, "包括以下应用下载域名\norigin -a.akamaihd.net"); 

            if (Properties.Settings.Default.NextUpdate == 0)
            {
                Properties.Settings.Default.NextUpdate = DateTime.Now.AddDays(7).Ticks;
                Properties.Settings.Default.Save();
            }
            else if (DateTime.Compare(DateTime.Now, new DateTime(Properties.Settings.Default.NextUpdate)) >= 0)
            {
                Properties.Settings.Default.NextUpdate = DateTime.Now.AddDays(7).Ticks;
                Properties.Settings.Default.Save();
                ThreadPool.QueueUserWorkItem(delegate { UpdateFile.Start(true); });
            }
        }

        private void TsmiMinimizeTray_Click(object sender, EventArgs e)
        {
            this.Hide();
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.ShowBalloonTip(1000);
        }

        private void TsmiExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void NotifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.Visible = true;
            this.notifyIcon1.Visible = false;
        }

        private void TsmUpdate_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.NextUpdate = DateTime.Now.AddDays(7).Ticks;
            Properties.Settings.Default.Save();
            ThreadPool.QueueUserWorkItem(delegate { UpdateFile.Start(false); });
        }

        private void TsmProductManual_Click(object sender, EventArgs e)
        {
            tsmProductManual.Enabled = false;
            FileInfo fi = new FileInfo(Application.StartupPath + "\\" + UpdateFile.pdfFile);
            if (!fi.Exists)
            {
                UpdateFile.bDownloadEnd = false;
                ThreadPool.QueueUserWorkItem(delegate { UpdateFile.Download(fi.Name); });
                while (!UpdateFile.bDownloadEnd)
                {
                    Application.DoEvents();
                }
                fi.Refresh();
            }
            if (fi.Exists)
                Process.Start(fi.FullName);
            else
                MessageBox.Show("文件不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            tsmProductManual.Enabled = true;
        }

        private void TsmForum_Click(object sender, EventArgs e)
        {
            ToolStripItem tsm = sender as ToolStripItem;
            Process.Start(tsm.Tag.ToString());
        }

        private void TsmAbout_Click(object sender, EventArgs e)
        {
            FormAbout dialog = new FormAbout();
            dialog.ShowDialog();
            dialog.Dispose();
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.M:
                    if (e.Control && e.Alt)
                    {
                        using (FileStream fs = File.Create(Application.ExecutablePath + ".md5"))
                        {
                            Byte[] b = new UTF8Encoding(true).GetBytes(UpdateFile.GetPathMD5(Application.ExecutablePath));
                            fs.Write(b, 0, b.Length);
                            fs.Close();
                        }
                    }
                    break;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (bServiceFlag) ButStart_Click(null, null);
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedTab.Name)
            {
                case "tabPage5":
                    if (Environment.OSVersion.Version.Major >= 10)
                    {
                        if (cbGameXGP1.Items.Count == 0 || cbGameXGP1.Items[0].ToString().Contains("(加载失败)") || cbGameXGP1.Items[cbGameXGP1.Items.Count - 1].ToString().Contains("(加载失败)"))
                        {
                            cbGameXGP1.Items.Clear();
                            cbGameXGP1.Items.Add(new Product("最受欢迎 Xbox Game Pass 游戏 (加载中)", "0"));
                            cbGameXGP1.SelectedIndex = 0;
                            ThreadPool.QueueUserWorkItem(delegate { GameXGPRecentlyAdded(1); });
                        }
                        if (cbGameXGP2.Items.Count == 0 || cbGameXGP2.Items[0].ToString().Contains("(加载失败)") || cbGameXGP2.Items[cbGameXGP2.Items.Count - 1].ToString().Contains("(加载失败)"))
                        {
                            cbGameXGP2.Items.Clear();
                            cbGameXGP2.Items.Add(new Product("近期新增 Xbox Game Pass 游戏 (加载中)", "0"));
                            cbGameXGP2.SelectedIndex = 0;
                            ThreadPool.QueueUserWorkItem(delegate { GameXGPRecentlyAdded(2); });
                        }
                    }
                    else if (cbGameXGP1.Items.Count == 0)
                    {
                        cbGameXGP1.Items.Add(new Product("最受欢迎 Xbox Game Pass 游戏 (不支持)", "0"));
                        cbGameXGP1.SelectedIndex = 0;
                        cbGameXGP2.Items.Add(new Product("近期新增 Xbox Game Pass 游戏 (不支持)", "0"));
                        cbGameXGP2.SelectedIndex = 0;
                    }
                    if (flpGameWithGold.Controls.Count == 0)
                    {
                        ThreadPool.QueueUserWorkItem(delegate { GameWithGold(); });
                    }
                    break;
                case "tabPage7":
                    if (cbAppxDrive.Items.Count == 0)
                    {
                        LinkAppxRefreshDrive_LinkClicked(null, null);
                    }
                    break;
            }
        }

        #region 选项卡-服务
        private void ButBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog
            {
                Description = "选择本地上传文件夹",
                SelectedPath = tbLocalPath.Text
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                tbLocalPath.Text = dlg.SelectedPath;
            }
        }

        private void ButStart_Click(object sender, EventArgs e)
        {
            if (bServiceFlag)
            {
                bServiceFlag = false;
                if (string.IsNullOrEmpty(Properties.Settings.Default.DnsIP)) tbDnsIP.Clear();
                if (string.IsNullOrEmpty(Properties.Settings.Default.ComIP)) tbComIP.Clear();
                if (string.IsNullOrEmpty(Properties.Settings.Default.CnIP)) tbCnIP.Clear();
                if (string.IsNullOrEmpty(Properties.Settings.Default.AppIP)) tbAppIP.Clear();
                if (string.IsNullOrEmpty(Properties.Settings.Default.EaIP)) tbEaIP.Clear();
                pictureBox1.Image = Properties.Resources.Xbox1;
                butStart.Text = "开始监听";
                tbDnsIP.Enabled = tbComIP.Enabled = tbCnIP.Enabled = tbAppIP.Enabled = tbEaIP.Enabled = ckbRedirect.Enabled = ckbTruncation.Enabled = ckbLocalUpload.Enabled = tbLocalPath.Enabled = butBrowse.Enabled = cbListenIP.Enabled = ckbDnsService.Enabled = ckbHttpService.Enabled =  ckbMicrosoftStore.Enabled = cbLocalIP.Enabled = true;
                if (Properties.Settings.Default.MicrosoftStore) ModifyHostsFile(false);
                linkTestDns.Enabled = false;
                dnsListen.Close();
                httpListen.Close();
                Program.SystemSleep.RestoreForCurrentThread();
            }
            else
            {
                string sRuleName = this.Text, sRulePath = Application.ExecutablePath;
                bool bRuleAdd = true;
                try
                {
                    INetFwPolicy2 policy2 = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
                    foreach (INetFwRule rule in policy2.Rules)
                    {
                        if (rule.Name == sRuleName)
                        {
                            if (bRuleAdd && rule.ApplicationName == sRulePath && rule.Direction == NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN && rule.Protocol == (int)NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_ANY && rule.Action == NET_FW_ACTION_.NET_FW_ACTION_ALLOW && rule.Profiles == (int)NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_ALL && rule.Enabled)
                                bRuleAdd = false;
                            else
                                policy2.Rules.Remove(rule.Name);
                        }
                        else if (String.Equals(rule.ApplicationName, sRulePath, StringComparison.CurrentCultureIgnoreCase))
                        {
                            policy2.Rules.Remove(rule.Name);
                        }
                    }
                    if (bRuleAdd)
                    {
                        INetFwRule rule = (INetFwRule)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwRule"));
                        rule.Name = sRuleName;
                        rule.ApplicationName = sRulePath;
                        rule.Enabled = true;
                        policy2.Rules.Add(rule);
                    }
                }
                catch { }

                string dnsIP = string.Empty;
                if (!string.IsNullOrEmpty(tbDnsIP.Text.Trim()))
                {
                    if (IPAddress.TryParse(tbDnsIP.Text, out IPAddress ipAddress))
                    {
                        dnsIP = ipAddress.ToString();
                    }
                    else
                    {
                        MessageBox.Show("DNS 服务器 IP 不正确", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tbDnsIP.Focus();
                        return;
                    }
                }
                string comIP = string.Empty;
                if (!string.IsNullOrEmpty(tbComIP.Text.Trim()))
                {
                    if (IPAddress.TryParse(tbComIP.Text, out IPAddress ipAddress))
                    {
                        comIP = ipAddress.ToString();
                    }
                    else
                    {
                        MessageBox.Show("指定 com 下载域名 IP 不正确", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tbComIP.Focus();
                        return;
                    }
                }
                string cnIP = string.Empty;
                if (!string.IsNullOrEmpty(tbCnIP.Text.Trim()))
                {
                    if (IPAddress.TryParse(tbCnIP.Text, out IPAddress ipAddress))
                    {
                        cnIP = ipAddress.ToString();
                    }
                    else
                    {
                        MessageBox.Show("指定 cn 下载域名 IP 不正确", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tbCnIP.Focus();
                        return;
                    }
                }
                string appIP = string.Empty;
                if (!string.IsNullOrEmpty(tbAppIP.Text.Trim()))
                {
                    if (IPAddress.TryParse(tbAppIP.Text, out IPAddress ipAddress))
                    {
                        appIP = ipAddress.ToString();
                    }
                    else
                    {
                        MessageBox.Show("指定应用下载域名 IP 不正确", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tbAppIP.Focus();
                        return;
                    }
                }
                string eaIP = string.Empty;
                if (!string.IsNullOrEmpty(tbEaIP.Text.Trim()))
                {
                    if (IPAddress.TryParse(tbEaIP.Text, out IPAddress ipAddress))
                    {
                        eaIP = ipAddress.ToString();
                    }
                    else
                    {
                        MessageBox.Show("指定 EA 下载域名 IP 不正确", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tbEaIP.Focus();
                        return;
                    }
                }
                Properties.Settings.Default.DnsIP = dnsIP;
                Properties.Settings.Default.ComIP = comIP;
                Properties.Settings.Default.CnIP = cnIP;
                Properties.Settings.Default.AppIP = appIP;
                Properties.Settings.Default.EaIP = eaIP;
                Properties.Settings.Default.Redirect = ckbRedirect.Checked;
                Properties.Settings.Default.Truncation = ckbTruncation.Checked;
                Properties.Settings.Default.LocalUpload = ckbLocalUpload.Checked;
                Properties.Settings.Default.LocalPath = tbLocalPath.Text;
                Properties.Settings.Default.ListenIP = cbListenIP.SelectedIndex;
                Properties.Settings.Default.DnsService = ckbDnsService.Checked;
                Properties.Settings.Default.HttpService = ckbHttpService.Checked;
                Properties.Settings.Default.MicrosoftStore = ckbMicrosoftStore.Checked;
                Properties.Settings.Default.Save();

                string resultInfo = string.Empty;
                using (Process p = new Process())
                {
                    p.StartInfo = new ProcessStartInfo("netstat", @"-aon")
                    {
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardOutput = true
                    };
                    p.Start();
                    resultInfo = p.StandardOutput.ReadToEnd();
                    p.Close();
                }
                Match result = Regex.Match(resultInfo, @"(?<protocol>TCP|UDP)\s+(?<ip>[^\s]+):(?<port>80|53)\s+[^\s]+\s+(?<status>[^\s]+\s+)?(?<pid>\d+)", RegexOptions.IgnoreCase);
                if (result.Success)
                {
                    ConcurrentDictionary<Int32, Process> dic = new ConcurrentDictionary<Int32, Process>();
                    StringBuilder sb = new StringBuilder();
                    while (result.Success)
                    {
                        if (Properties.Settings.Default.ListenIP == 0)
                        {
                            if (result.Groups["ip"].Value == Properties.Settings.Default.LocalIP)
                            {
                                string protocol = result.Groups["protocol"].Value;
                                if (protocol == "TCP" && result.Groups["status"].Value == "LISTENING" || protocol == "UDP")
                                {
                                    int port = Convert.ToInt32(result.Groups["port"].Value);
                                    if (port == 53 && Properties.Settings.Default.DnsService || port == 80 && Properties.Settings.Default.HttpService)
                                    {
                                        int pid = int.Parse(result.Groups["pid"].Value);
                                        if (!dic.ContainsKey(pid) && pid != 0)
                                        {
                                            sb.AppendLine(protocol + "\t" + result.Groups["ip"].Value + ":" + port);
                                            if (pid == 4)
                                            {
                                                dic.TryAdd(pid, null);
                                                sb.AppendLine("System");
                                            }
                                            else
                                            {
                                                try
                                                {
                                                    Process procId = Process.GetProcessById(pid);
                                                    dic.TryAdd(pid, procId);
                                                    string filename = procId.MainModule.FileName;
                                                    sb.AppendLine(filename);
                                                }
                                                catch
                                                {
                                                    sb.AppendLine("未知");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            string protocol = result.Groups["protocol"].Value;
                            int port = Convert.ToInt32(result.Groups["port"].Value);
                            if (port == 53 && Properties.Settings.Default.DnsService || port == 80 && Properties.Settings.Default.HttpService)
                            {
                                int pid = int.Parse(result.Groups["pid"].Value);
                                if (!dic.ContainsKey(pid) && pid != 0)
                                {
                                    sb.AppendLine(protocol + "\t" + result.Groups["ip"].Value + ":" + port);
                                    if (pid == 4)
                                    {
                                        dic.TryAdd(pid, null);
                                        sb.AppendLine("System");
                                    }
                                    else
                                    {
                                        try
                                        {
                                            Process procId = Process.GetProcessById(pid);
                                            dic.TryAdd(pid, procId);
                                            string filename = procId.MainModule.FileName;
                                            sb.AppendLine(filename);
                                        }
                                        catch
                                        {
                                            sb.AppendLine("未知");
                                        }
                                    }
                                }
                            }
                        }
                        result = result.NextMatch();
                    }
                    if (dic.Count >= 1 && MessageBox.Show("检测到以下端口被占用\n" + sb.ToString() + "\n是否尝试强制结束占用端口程序？", "启用服务失败", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        foreach (var item in dic)
                        {
                            if (item.Key == 4)
                            {
                                using (Process p = new Process())
                                {
                                    p.StartInfo.FileName = "cmd.exe";
                                    p.StartInfo.UseShellExecute = false;
                                    p.StartInfo.RedirectStandardInput = true;
                                    p.StartInfo.RedirectStandardError = true;
                                    p.StartInfo.CreateNoWindow = true;
                                    p.Start();

                                    p.StandardInput.WriteLine("iisreset /stop");
                                    p.StandardInput.WriteLine("net stop \"SQL Server Reporting Services (MSSQLSERVER)\" /Y");
                                    p.StandardInput.WriteLine("exit");

                                    p.WaitForExit();
                                    p.Close();
                                }
                            }
                            else
                            {
                                try
                                {
                                    item.Value.Kill();
                                }
                                catch { }
                            }
                        }
                    }
                }

                bServiceFlag = true;
                pictureBox1.Image = Properties.Resources.Xbox2;
                tbDnsIP.Enabled = tbComIP.Enabled = tbCnIP.Enabled = tbAppIP.Enabled = tbEaIP.Enabled = ckbRedirect.Enabled = ckbTruncation.Enabled = ckbLocalUpload.Enabled = tbLocalPath.Enabled = butBrowse.Enabled = cbListenIP.Enabled = ckbDnsService.Enabled = ckbHttpService.Enabled = ckbMicrosoftStore.Enabled = cbLocalIP.Enabled = false;
                butStart.Text = "停止监听";
                Program.SystemSleep.PreventForCurrentThread(false);

                if (Properties.Settings.Default.DnsService)
                {
                    linkTestDns.Enabled = true;
                    string[] ips = Properties.Settings.Default.LocalIP.Split('.');
                    Byte[] ipByte = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
                    dicDomain.AddOrUpdate(Environment.MachineName, ipByte, (oldkey, oldvalue) => ipByte);

                    Thread dnsThread = new Thread(new ThreadStart(dnsListen.Listen))
                    {
                        IsBackground = true
                    };
                    dnsThread.Start();
                }
                if (Properties.Settings.Default.HttpService)
                {
                    Thread httpThread = new Thread(new ThreadStart(httpListen.Listen))
                    {
                        IsBackground = true
                    };
                    httpThread.Start();
                }
                if (Properties.Settings.Default.MicrosoftStore)
                {
                    ModifyHostsFile(true);
                }
            }
        }

        private void ModifyHostsFile(bool add)
        {
            string sHostsPath = Environment.SystemDirectory + "\\drivers\\etc\\hosts";
            try
            {
                FileInfo fi = new FileInfo(sHostsPath);
                if (!fi.Exists)
                {
                    StreamWriter sw = fi.CreateText();
                    sw.Close();
                    fi.Refresh();
                }
                if ((fi.Attributes & FileAttributes.ReadOnly) != 0)
                    fi.Attributes = FileAttributes.Normal;
                FileSecurity fSecurity = fi.GetAccessControl();
                fSecurity.AddAccessRule(new FileSystemAccessRule("Administrators", FileSystemRights.FullControl, AccessControlType.Allow));
                fi.SetAccessControl(fSecurity);
                string sHosts = string.Empty;
                using (StreamReader sw = new StreamReader(sHostsPath))
                {
                    sHosts = sw.ReadToEnd();
                }
                sHosts = Regex.Replace(sHosts, @"# Added by Xbox下载助手\r\n(.*\r\n)+# End of Xbox下载助手\r\n", "");
                if (add)
                {
                    string comIP = string.IsNullOrEmpty(Properties.Settings.Default.ComIP) ? Properties.Settings.Default.LocalIP : Properties.Settings.Default.ComIP;
                    if (!Properties.Settings.Default.DnsService && string.IsNullOrEmpty(Properties.Settings.Default.ComIP))
                        tbComIP.Text = comIP;
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("# Added by Xbox下载助手");
                    sb.AppendLine(comIP + " assets1.xboxlive.com");
                    sb.AppendLine(comIP + " assets2.xboxlive.com");
                    sb.AppendLine(comIP + " dlassets.xboxlive.com");
                    sb.AppendLine(comIP + " dlassets2.xboxlive.com");
                    sb.AppendLine(comIP + " d1.xboxlive.com");
                    sb.AppendLine(comIP + " d2.xboxlive.com");
                    sb.AppendLine(comIP + " xvcf1.xboxlive.com");
                    sb.AppendLine(comIP + " xvcf2.xboxlive.com");
                    if (!string.IsNullOrEmpty(Properties.Settings.Default.CnIP))
                    {
                        sb.AppendLine(Properties.Settings.Default.CnIP + " assets1.xboxlive.cn");
                        sb.AppendLine(Properties.Settings.Default.CnIP + " assets2.xboxlive.cn");
                        sb.AppendLine(Properties.Settings.Default.CnIP + " dlassets.xboxlive.cn");
                        sb.AppendLine(Properties.Settings.Default.CnIP + " dlassets2.xboxlive.cn");
                        sb.AppendLine(Properties.Settings.Default.CnIP + " d1.xboxlive.cn");
                        sb.AppendLine(Properties.Settings.Default.CnIP + " d2.xboxlive.cn");
                    }
                    if (!string.IsNullOrEmpty(Properties.Settings.Default.AppIP))
                    {
                        sb.AppendLine(Properties.Settings.Default.AppIP + " dl.delivery.mp.microsoft.com");
                        sb.AppendLine(Properties.Settings.Default.AppIP + " tlu.dl.delivery.mp.microsoft.com");
                    }
                    if (!string.IsNullOrEmpty(Properties.Settings.Default.EaIP))
                    {
                        sb.AppendLine(Properties.Settings.Default.EaIP + " origin-a.akamaihd.net");
                    }
                    foreach (var domain in dicDomain)
                    {
                        if (domain.Key == Environment.MachineName)
                            continue;
                        sb.AppendLine(string.Format("{0}.{1}.{2}.{3} {4}", domain.Value[0], domain.Value[1], domain.Value[2], domain.Value[3], domain.Key));
                    }
                    sb.AppendLine("# End of Xbox下载助手");
                    sHosts += sb.ToString();
                }
                using (StreamWriter sw = new StreamWriter(sHostsPath, false))
                {
                    sw.Write(sHosts.Trim() + "\r\n");
                }
                fSecurity.RemoveAccessRule(new FileSystemAccessRule("Administrators", FileSystemRights.FullControl, AccessControlType.Allow));
                fi.SetAccessControl(fSecurity);
            }
            catch (Exception e)
            {
                if (add) MessageBox.Show("加速应用商店(PC)失败，错误信息：" + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LvLog_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (lvLog.SelectedItems.Count == 1)
                {
                    tsmCopy.Visible = true;
                    tsmUseIP.Visible = false;
                    contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
                }
            }
        }

        private void CbLocalIP_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.LocalIP = cbLocalIP.Text;
            Properties.Settings.Default.Save();
        }

        private void LinkTestDns_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormDns dialog = new FormDns();
            dialog.ShowDialog();
            dialog.Dispose();
        }

        private void CbRecordLog_CheckedChanged(object sender, EventArgs e)
        {
            bRecordLog = ckbRecordLog.Checked;
        }

        private void LinkClearLog_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lvLog.Items.Clear();
        }
        #endregion

        #region 选项卡-测速
        private void DgvIpList_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || e.Button != MouseButtons.Right) return;
            string host = string.Empty;
            Match result = Regex.Match(groupBox4.Text, @"\((?<host>[^\)]+)\)");
            if (result.Success) host = result.Groups["host"].Value;
            dgvIpList.ClearSelection();
            DataGridViewRow dgvr = dgvIpList.Rows[e.RowIndex];
            dgvr.Selected = true;
            tsmCopy.Visible = false;
            tsmUseIP.Visible = true;
            if (host == "origin-a.akamaihd.net")
            {
                tsmUseIP1.Visible = tsmUseIP2.Visible = tsmUseIP3.Visible = false;
                tsmUseIP4.Visible = true;
            }
            else
            {
                tsmUseIP1.Text = (Regex.IsMatch(host, @"\.xboxlive\.com")) ? "指定 com 下载域名 IP" : "指定 cn 下载域名 IP";
                tsmUseIP1.Visible = tsmUseIP2.Visible = tsmUseIP3.Visible = true;
                tsmUseIP4.Visible = false;
            }
            contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
        }

        private void TsmCopy_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(lvLog.SelectedItems[0].SubItems[1].Text);
        }

        private void TsmUseIP_Click(object sender, EventArgs e)
        {
            if (bServiceFlag)
            {
                MessageBox.Show("请先停止监听后再设置。", "使用指定IP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dgvIpList.SelectedRows.Count != 1) return;
            DataGridViewRow dgvr = dgvIpList.SelectedRows[0];
            string ip = dgvr.Cells["Col_IP"].Value.ToString();
            ToolStripMenuItem tsmi = sender as ToolStripMenuItem;
            switch (tsmi.Name)
            {
                case "tsmUseIP1":
                    if (tsmUseIP1.Text == "指定 com 下载域名 IP")
                    {
                        tbComIP.Text = ip;
                        tabControl1.SelectedIndex = 0;
                        tbComIP.Focus();
                    }
                    else
                    {
                        tbCnIP.Text = ip;
                        tabControl1.SelectedIndex = 0;
                        tbCnIP.Focus();
                    }
                    break;
                case "tsmUseIP2":
                    tbAppIP.Text = ip;
                    tabControl1.SelectedIndex = 0;
                    tbAppIP.Focus();
                    break;
                case "tsmUseIP3":
                    if (tsmUseIP1.Text == "指定 com 下载域名 IP")
                    {
                        tbComIP.Text = ip;
                        tabControl1.SelectedIndex = 0;
                        tbComIP.Focus();
                    }
                    else
                    {
                        tbCnIP.Text = ip;
                        tabControl1.SelectedIndex = 0;
                        tbCnIP.Focus();
                    }
                    tbAppIP.Text = ip;
                    break;
            }
        }

        private void TsmUseIP4_Click(object sender, EventArgs e)
        {
            if (bServiceFlag)
            {
                MessageBox.Show("请先停止监听后再设置。", "使用指定IP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dgvIpList.SelectedRows.Count != 1) return;
            DataGridViewRow dgvr = dgvIpList.SelectedRows[0];
            tbEaIP.Text = dgvr.Cells["Col_IP"].Value.ToString();
            tabControl1.SelectedIndex = 0;
            tbEaIP.Focus();
        }

        private void DgvIpList_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == -1) return;
            if (e.Button == MouseButtons.Left && dgvIpList.Columns[dgvIpList.CurrentCell.ColumnIndex].Name == "Col_Speed" && dgvIpList.Rows[e.RowIndex].Tag != null)
            {
                string msg = dgvIpList.Rows[e.RowIndex].Tag.ToString().Trim();
                if (!string.IsNullOrEmpty(msg))
                    MessageBox.Show(msg, "Request Headers", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }

        private void CkbTelecom_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = (CheckBox)sender;
            string telecom = cb.Text;
            bool isChecked = cb.Checked;
            foreach (DataGridViewRow dgvr in dgvIpList.Rows)
            {
                string ASN = dgvr.Cells["Col_ASN"].Value.ToString();
                if (telecom == "其它")
                {
                    if (!Regex.IsMatch(ASN, @"电信|联通|移动") || ASN.Contains("中华电信"))
                        dgvr.Cells["Col_Check"].Value = isChecked;
                }
                else
                {
                    if (ASN.Contains(telecom) && !ASN.Contains("中华电信"))
                        dgvr.Cells["Col_Check"].Value = isChecked;
                }
            }
        }

        private void LinkTestUrl_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel link = sender == null ? linkTestUrl1 : sender as LinkLabel;
            string url = link.Tag.ToString();
            Match result = Regex.Match(groupBox4.Text, @"\((?<host>[^\)]+)\)");
            if (result.Success) url = "http://" + result.Groups["host"].Value + url;
            tbDlUrl.Text = url;
        }

        private void LinkImportIP_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel link = sender as LinkLabel;
            bool xbox = link.Name == "linkImportIPXbox";
            dgvIpList.Tag = xbox ? "assets1.xboxlive.cn" : "origin-a.akamaihd.net";

            dgvIpList.Rows.Clear();
            tbDlUrl.Clear();
            linkImportIPXbox.Enabled = linkImportIPEA.Enabled = false;
            if (!xbox) linkTestUrl1.Enabled = linkTestUrl2.Enabled = linkTestUrl3.Enabled = false;

            bool update = true;
            FileInfo fi = new FileInfo(Application.StartupPath + "\\IP." + dgvIpList.Tag + ".txt");
            if (fi.Exists) update = DateTime.Compare(DateTime.Now, fi.LastWriteTime.AddHours(24)) >= 0;
            if (update)
            {
                UpdateFile.bDownloadEnd = false;
                ThreadPool.QueueUserWorkItem(delegate { UpdateFile.Download(fi.Name); });
                while (!UpdateFile.bDownloadEnd)
                {
                    Application.DoEvents();
                }
                fi.Refresh();
            }
            string content = string.Empty;
            if (fi.Exists)
            {
                using (StreamReader sr = fi.OpenText())
                {
                    content = sr.ReadToEnd();
                }
            }

            List<DataGridViewRow> list = new List<DataGridViewRow>();
            bool telecom1 = ckbTelecom1.Checked;
            bool telecom2 = ckbTelecom2.Checked;
            bool telecom3 = ckbTelecom3.Checked;
            bool telecom4 = ckbTelecom4.Checked;
            Match result = Regex.Match(content, @"(?<IP>\d{0,3}\.\d{0,3}\.\d{0,3}\.\d{0,3})\s*\((?<ASN>[^\)]+)\)|(?<IP>\d{0,3}\.\d{0,3}\.\d{0,3}\.\d{0,3})(?<ASN>.+)\dms|^\s*(?<IP>\d{0,3}\.\d{0,3}\.\d{0,3}\.\d{0,3})\s*$", RegexOptions.Multiline);
            if (result.Success)
            {
                groupBox4.Text = "IP 列表 (" + dgvIpList.Tag + ")";
                while (result.Success)
                {
                    string ip = result.Groups["IP"].Value;
                    string ASN = result.Groups["ASN"].Value.Trim();

                    DataGridViewRow dgvr = new DataGridViewRow();
                    dgvr.CreateCells(dgvIpList);
                    dgvr.Resizable = DataGridViewTriState.False;
                    if (telecom1 && ASN.Contains("电信") && !ASN.Contains("中华电信") || telecom2 && ASN.Contains("联通") || telecom3 && ASN.Contains("移动") || (telecom4 && (!Regex.IsMatch(ASN, @"电信|联通|移动") || ASN.Contains("中华电信"))))
                        dgvr.Cells[0].Value = true;
                    dgvr.Cells[1].Value = ip;
                    dgvr.Cells[2].Value = ASN;
                    list.Add(dgvr);
                    result = result.NextMatch();
                }
                if (list.Count >= 1)
                {
                    dgvIpList.Rows.AddRange(list.ToArray());
                    dgvIpList.ClearSelection();
                }
            }
            if (xbox) linkTestUrl1.Enabled = linkTestUrl2.Enabled = linkTestUrl3.Enabled = true;
            linkImportIPXbox.Enabled = linkImportIPEA.Enabled = true;
        }

        private void LinkImportIPManual_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormImportIP dialog = new FormImportIP();
            dialog.ShowDialog();
            string host = dialog.host;
            DataTable dt = dialog.dt;
            dialog.Dispose();
            if (dt != null && dt.Rows.Count >= 1)
            {
                dgvIpList.Rows.Clear();
                dgvIpList.Tag = host;
                bool telecom1 = ckbTelecom1.Checked;
                bool telecom2 = ckbTelecom2.Checked;
                bool telecom3 = ckbTelecom3.Checked;
                bool telecom4 = ckbTelecom4.Checked;
                List<DataGridViewRow> list = new List<DataGridViewRow>();
                groupBox4.Text = "IP 列表 (" + dgvIpList.Tag + ")";
                foreach (DataRow dr in dt.Select("", "ASN, IpLong"))
                {
                    string ASN = dr["ASN"].ToString();
                    DataGridViewRow dgvr = new DataGridViewRow();
                    dgvr.CreateCells(dgvIpList);
                    dgvr.Resizable = DataGridViewTriState.False;
                    if (telecom1 && ASN.Contains("电信") && !ASN.Contains("中华电信") || telecom2 && ASN.Contains("联通") || telecom3 && ASN.Contains("移动") || (telecom4 && (!Regex.IsMatch(ASN, @"电信|联通|移动") || ASN.Contains("中华电信"))))
                        dgvr.Cells[0].Value = true;
                    dgvr.Cells[1].Value = dr["IP"];
                    dgvr.Cells[2].Value = dr["ASN"];
                    list.Add(dgvr);
                }
                if (list.Count >= 1)
                {
                    dgvIpList.Rows.AddRange(list.ToArray());
                    dgvIpList.ClearSelection();
                }
            }
        }

        private void LinkExportIP_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (dgvIpList.Rows.Count == 0) return;
            string host = dgvIpList.Tag.ToString();
            SaveFileDialog dlg = new SaveFileDialog
            {
                InitialDirectory = Application.StartupPath,
                Title = "导出数据",
                Filter = "文本文件(*.txt)|*.txt",
                FileName = "导出IP(" + host + ")"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(host);
                sb.AppendLine("");
                foreach (DataGridViewRow dgvr in dgvIpList.Rows)
                {
                    if (dgvr.Cells["Col_Speed"].Value != null && !string.IsNullOrEmpty(dgvr.Cells["Col_Speed"].Value.ToString()))
                        sb.AppendLine(dgvr.Cells["Col_IP"].Value + "\t(" + dgvr.Cells["Col_ASN"].Value + ")\t" + dgvr.Cells["Col_Speed"].Value + "Mbps");
                    else
                        sb.AppendLine(dgvr.Cells["Col_IP"].Value + "\t(" + dgvr.Cells["Col_ASN"].Value + ")");
                }
                using (FileStream fs = File.Create(dlg.FileName))
                {
                    Byte[] log = new UTF8Encoding(true).GetBytes(sb.ToString());
                    fs.Write(log, 0, log.Length);
                    fs.Close();
                }
            }
        }

        bool isSpeedTest = false;
        Thread threadSpeedTest = null;
        private void ButSpeedTest_Click(object sender, EventArgs e)
        {
            if (!isSpeedTest)
            {
                if (dgvIpList.Rows.Count == 0)
                {
                    MessageBox.Show("请先导入IP。", "IP列表为空", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                List<DataGridViewRow> ls = new List<DataGridViewRow>();
                foreach (DataGridViewRow dgvr in dgvIpList.Rows)
                {
                    if (Convert.ToBoolean(dgvr.Cells["Col_Check"].Value))
                    {
                        ls.Add(dgvr);
                    }
                }
                if (ls.Count == 0)
                {
                    MessageBox.Show("请勾选需要测试IP。", "选择测试IP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int rowIndex = 0;
                foreach (DataGridViewRow dgvr in ls.ToArray())
                {
                    dgvIpList.Rows.Remove(dgvr);
                    dgvIpList.Rows.Insert(rowIndex, dgvr);
                    rowIndex++;
                }
                dgvIpList.Rows[0].Cells[0].Selected = true;

                string host = dgvIpList.Tag.ToString();
                string dlFile = tbDlUrl.Text.Trim();
                if (string.IsNullOrEmpty(dlFile))
                {
                    if (host == "origin-a.akamaihd.net")
                        tbDlUrl.Text = "https://origin-a.akamaihd.net/EA-Desktop-Client-Download/installer-releases/EADesktop-12.0.84.4906-347.msi";
                    else
                        LinkTestUrl_LinkClicked(null, null);
                    dlFile = tbDlUrl.Text;
                }
                if (!Regex.IsMatch(dlFile, @"https?://"))
                {
                    dlFile = "http://" + host + dlFile;
                    tbDlUrl.Text = dlFile;
                }
                isSpeedTest = true;
                butSpeedTest.Text = "停止测速";
                ckbTelecom1.Enabled = ckbTelecom2.Enabled = ckbTelecom3.Enabled = ckbTelecom4.Enabled = linkExportIP.Enabled = linkImportIPXbox.Enabled = linkImportIPEA.Enabled = linkImportIPManual.Enabled = tbDlUrl.Enabled = linkTestUrl1.Enabled = linkTestUrl2.Enabled = linkTestUrl3.Enabled = false;
                Col_IP.SortMode = Col_ASN.SortMode = Col_TTL.SortMode = Col_RoundtripTime.SortMode = Col_Speed.SortMode = DataGridViewColumnSortMode.NotSortable;
                Col_Check.ReadOnly = true;
                threadSpeedTest = new Thread(new ThreadStart(() =>
                {
                    SpeedTest(ls, dlFile);
                }))
                {
                    IsBackground = true
                };
                threadSpeedTest.Start();
            }
            else
            {
                if (threadSpeedTest != null && threadSpeedTest.IsAlive) threadSpeedTest.Abort();
                foreach (DataGridViewRow dgvr in dgvIpList.Rows)
                {
                    if (dgvr.Cells["Col_Speed"].Value != null && dgvr.Cells["Col_Speed"].Value.ToString() == "正在测试")
                    {
                        dgvr.Cells["Col_Speed"].Value = null;
                        break;
                    }
                }
                butSpeedTest.Text = "开始测速";
                ckbTelecom1.Enabled = ckbTelecom2.Enabled = ckbTelecom3.Enabled = ckbTelecom4.Enabled = linkExportIP.Enabled = linkImportIPXbox.Enabled = linkImportIPEA.Enabled = linkImportIPManual.Enabled = tbDlUrl.Enabled = true;
                if (dgvIpList.Tag.ToString() != "origin-a.akamaihd.net") linkTestUrl1.Enabled = linkTestUrl2.Enabled = linkTestUrl3.Enabled = true;
                isSpeedTest = false;
            }
        }

        private void SpeedTest(List<DataGridViewRow> ls, string dlFile)
        {
            string[] headers = new string[] { "Range: bytes=0-209715199" }; //200M
            //string[] headers = new string[] { "Range: bytes=0-1048575" }; //1M
            Stopwatch sw = new Stopwatch();
            foreach (DataGridViewRow dgvr in ls)
            {
                string ip = dgvr.Cells["Col_IP"].Value.ToString();
                dgvr.Cells["Col_TTL"].Value = null;
                dgvr.Cells["Col_RoundtripTime"].Value = null;
                dgvr.Cells["Col_Speed"].Value = "正在测试";
                dgvr.Cells["Col_Speed"].Style.ForeColor = Color.Empty;
                dgvr.Tag = null;

                using (Ping p1 = new Ping())
                {
                    try
                    {
                        PingReply reply = p1.Send(ip);
                        if (reply.Status == IPStatus.Success)
                        {
                            dgvr.Cells["Col_TTL"].Value = reply.Options.Ttl;
                            dgvr.Cells["Col_RoundtripTime"].Value = reply.RoundtripTime;
                        }
                    }
                    catch { }
                }
                sw.Restart();
                SocketPackage socketPackage = ClassWeb.HttpRequest(dlFile, "GET", null, null, true, false, false, null, null, headers, null, null, null, null, null, 0, null, 15000, 15000, 1, ip, true);
                sw.Stop();
                dgvr.Tag = string.IsNullOrEmpty(socketPackage.Err) ? socketPackage.Headers : socketPackage.Err;
                if (socketPackage.Headers.StartsWith("HTTP/1.1 206"))
                {
                    dgvr.Cells["Col_Speed"].Value = Math.Round((double)(socketPackage.Buffer.Length) / sw.ElapsedMilliseconds * 1000 * 8 / 1024 / 1024, 2, MidpointRounding.AwayFromZero);
                }
                else
                {
                    dgvr.Cells["Col_Speed"].Value = (double)0;
                    dgvr.Cells["Col_Speed"].Style.ForeColor = Color.Red;
                }
            }
            this.Invoke(new Action(() =>
            {
                butSpeedTest.Text = "开始测速";
                ckbTelecom1.Enabled = ckbTelecom2.Enabled = ckbTelecom3.Enabled = ckbTelecom4.Enabled = linkExportIP.Enabled = linkImportIPXbox.Enabled = linkImportIPEA.Enabled = linkImportIPManual.Enabled = tbDlUrl.Enabled = true;
                if (dgvIpList.Tag.ToString() != "origin-a.akamaihd.net") linkTestUrl1.Enabled = linkTestUrl2.Enabled = linkTestUrl3.Enabled = true;
                Col_IP.SortMode = Col_ASN.SortMode = Col_Speed.SortMode = Col_TTL.SortMode = Col_RoundtripTime.SortMode = DataGridViewColumnSortMode.Automatic;
                Col_Check.ReadOnly = false;
            }));
            isSpeedTest = false;
        }
        #endregion

        #region 选项卡-域名
        private void DgvHosts_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["Col_Enable"].Value = true;
        }

        private void LinkXbox360DomainName_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string[] dn = new string[] { "download.xbox.com", "download.xbox.com.edgesuite.net", "xbox-ecn102.vo.msecnd.net" };
            foreach (string domain in dn)
            {
                DataRow[] rows = dtDomain.Select("Domain='" + domain + "'");
                if (rows.Length >= 1)
                {
                    rows[0]["Enable"] = true;
                    rows[0]["IPv4"] = Properties.Settings.Default.LocalIP;
                    rows[0]["Remark"] = "Xbox360主机下载域名";
                }
                else
                {
                    DataRow dr = dtDomain.NewRow();
                    dr["Enable"] = true;
                    dr["Domain"] = domain;
                    dr["IPv4"] = Properties.Settings.Default.LocalIP;
                    dr["Remark"] = "Xbox360主机下载域名";
                    dtDomain.Rows.Add(dr);
                }
            }
        }

        private void ButDomainSave_Click(object sender, EventArgs e)
        {
            dtDomain.AcceptChanges();
            if (dtDomain.Rows.Count >= 1)
                dtDomain.WriteXml(domainPath);
            else if (File.Exists(domainPath))
                File.Delete(domainPath);
            AddDomain();

            if (bServiceFlag && Properties.Settings.Default.MicrosoftStore)
            {
                ModifyHostsFile(true);
            }
        }

        private void ButDomainReset_Click(object sender, EventArgs e)
        {
            dtDomain.RejectChanges();
        }

        private void LinkDomainClear_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            for (int i = dgvHosts.Rows.Count - 2; i >= 0; i--)
            {
                dgvHosts.Rows.RemoveAt(i);
            }
        }

        private void AddDomain()
        {
            dicDomain.Clear();
            foreach (DataRow dr in dtDomain.Rows)
            {
                string domain = dr["Domain"].ToString().Trim().ToLower();
                if (Convert.ToBoolean(dr["Enable"]) && !string.IsNullOrEmpty(domain))
                {
                    if (IPAddress.TryParse(dr["IPv4"].ToString(), out IPAddress ip))
                    {
                        string[] ips = ip.ToString().Split('.');
                        Byte[] ipByte = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
                        dicDomain.TryAdd(domain, ipByte);
                    }
                }
            }
        }
        #endregion

        #region 选项卡-硬盘
        private void ButScan_Click(object sender, EventArgs e)
        {
            dgvDevice.Rows.Clear();
            butEnabelPc.Enabled = butEnabelXbox.Enabled = false;
            List<DataGridViewRow> list = new List<DataGridViewRow>();

            ManagementClass mc = new ManagementClass("Win32_DiskDrive");
            ManagementObjectCollection moc = mc.GetInstances();
            foreach (ManagementObject mo in moc)
            {
                string sDeviceID = mo.Properties["DeviceID"].Value.ToString();
                string mbr = ClassMbr.ByteToHexString(ClassMbr.ReadMBR(sDeviceID));
                if (string.Equals(mbr.Substring(0, 892), ClassMbr.MBR))
                {
                    string mode = mbr.Substring(1020);
                    DataGridViewRow dgvr = new DataGridViewRow();
                    dgvr.CreateCells(dgvDevice);
                    dgvr.Resizable = DataGridViewTriState.False;
                    dgvr.Tag = mode;
                    dgvr.Cells[0].Value = sDeviceID;
                    dgvr.Cells[1].Value = mo.Properties["Model"].Value;
                    dgvr.Cells[2].Value = mo.Properties["InterfaceType"].Value;
                    dgvr.Cells[3].Value = ClassMbr.ConvertBytes(Convert.ToUInt64(mo.Properties["Size"].Value));
                    if (mode == "99CC")
                        dgvr.Cells[4].Value = "Xbox 模式";
                    else if (mode == "55AA")
                        dgvr.Cells[4].Value = "PC 模式";
                    list.Add(dgvr);
                }
            }
            if (list.Count >= 1)
            {
                dgvDevice.Rows.AddRange(list.ToArray());
                dgvDevice.ClearSelection();
            }
        }

        private void DgvDevice_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1) return;
            string mode = dgvDevice.Rows[e.RowIndex].Tag.ToString();
            if (mode == "99CC")
            {
                butEnabelPc.Enabled = true;
                butEnabelXbox.Enabled = false;
            }
            else if (mode == "55AA")
            {
                butEnabelPc.Enabled = false;
                butEnabelXbox.Enabled = true;
            }
        }

        private void ButEnabelPc_Click(object sender, EventArgs e)
        {
            if (dgvDevice.SelectedRows.Count != 1) return;
            if (Environment.OSVersion.Version.Major < 10)
            {
                MessageBox.Show("低于Win10操作系统转换后会蓝屏，请升级操作系统。", "操作系统版本过低", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            string sDeviceID = dgvDevice.SelectedRows[0].Cells["Col_DeviceID"].Value.ToString();
            string mode = dgvDevice.SelectedRows[0].Tag.ToString();
            string mbr = ClassMbr.ByteToHexString(ClassMbr.ReadMBR(sDeviceID));
            if (mode == "99CC" && mbr.Substring(0, 892) == ClassMbr.MBR && mbr.Substring(1020) == mode)
            {
                string newMBR = mbr.Substring(0, 1020) + "55AA";
                if (ClassMbr.WriteMBR(sDeviceID, ClassMbr.HexToByte(newMBR)))
                {
                    dgvDevice.SelectedRows[0].Tag = "55AA";
                    dgvDevice.SelectedRows[0].Cells["Col_Mode"].Value = "PC 模式";
                    dgvDevice.ClearSelection();
                    butEnabelPc.Enabled = false;
                    using (Process p = new Process())
                    {
                        p.StartInfo.FileName = "diskpart.exe";
                        p.StartInfo.RedirectStandardInput = true;
                        p.StartInfo.CreateNoWindow = true;
                        p.StartInfo.UseShellExecute = false;
                        p.Start();
                        p.StandardInput.WriteLine("rescan");
                        p.StandardInput.WriteLine("exit");
                        p.Close();
                    }
                    MessageBox.Show("成功转换PC模式。", "转换PC模式", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void ButEnabelXbox_Click(object sender, EventArgs e)
        {
            if (dgvDevice.SelectedRows.Count != 1) return;
            string sDeviceID = dgvDevice.SelectedRows[0].Cells["Col_DeviceID"].Value.ToString();
            string mode = dgvDevice.SelectedRows[0].Tag.ToString();
            string mbr = ClassMbr.ByteToHexString(ClassMbr.ReadMBR(sDeviceID));
            if (mode == "55AA" && mbr.Substring(0, 892) == ClassMbr.MBR && mbr.Substring(1020) == mode)
            {
                string newMBR = mbr.Substring(0, 1020) + "99CC";
                if (ClassMbr.WriteMBR(sDeviceID, ClassMbr.HexToByte(newMBR)))
                {
                    dgvDevice.SelectedRows[0].Tag = "99CC";
                    dgvDevice.SelectedRows[0].Cells["Col_Mode"].Value = "Xbox 模式";
                    dgvDevice.ClearSelection();
                    butEnabelXbox.Enabled = false;
                    MessageBox.Show("成功转换Xbox模式。", "转换Xbox模式", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void Tb_Enter(object sender, EventArgs e)
        {
            BeginInvoke((Action)delegate
            {
                (sender as TextBox).SelectAll();
            });
        }

        private void ButDownload_Click(object sender, EventArgs e)
        {
            string url = tbDownloadUrl.Text.Trim();
            if (string.IsNullOrEmpty(url)) return;
            if (!Regex.IsMatch(url, @"^https?\:\/\/"))
            {
                if (!url.StartsWith("/")) url = "/" + url;
                url = "http://assets1.xboxlive.cn" + url;
                tbDownloadUrl.Text = url;
            }

            tbFilePath.Text = string.Empty;
            byte[] bFileBuffer = null;
            SocketPackage socketPackage = ClassWeb.HttpRequest(url, "GET", null, null, true, false, false, null, null, new string[] { "Range: bytes=0-4095" }, null, null, null, null, null, 0, null);
            if (!string.IsNullOrEmpty(socketPackage.Err))
            {
                MessageBox.Show("下载失败，错误信息：" + socketPackage.Err, "下载失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                bFileBuffer = socketPackage.Buffer;
            }
            Analysis(bFileBuffer);
        }

        private void ButOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Open an Xbox Package"
            };
            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string sFilePath = ofd.FileName;
            tbDownloadUrl.Text = "";
            tbFilePath.Text = sFilePath;

            FileStream fs = null;
            try
            {
                fs = new FileStream(sFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (fs != null)
            {
                int len = fs.Length >= 49152 ? 49152 : (int)fs.Length;
                byte[] bFileBuffer = new byte[len];
                fs.Read(bFileBuffer, 0, len);
                fs.Close();
                Analysis(bFileBuffer);
            }
        }

        private void Analysis(byte[] bFileBuffer)
        {
            tbContentId.Text = tbProductID.Text = tbBuildID.Text = tbFileTimeCreated.Text = tbDriveSize.Text = tbPackageVersion.Text = string.Empty;
            linkCopyContentID.Enabled = linkRename.Enabled = false;
            if (bFileBuffer.Length >= 4096)
            {
                using (MemoryStream ms = new MemoryStream(bFileBuffer))
                {
                    BinaryReader br = null;
                    try
                    {
                        br = new BinaryReader(ms);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    if (br != null)
                    {
                        br.BaseStream.Position = 0x200;
                        if (Encoding.Default.GetString(br.ReadBytes(0x8)) == "msft-xvd")
                        {
                            br.BaseStream.Position = 0x210;
                            tbFileTimeCreated.Text = DateTime.FromFileTime(BitConverter.ToInt64(br.ReadBytes(0x8), 0)).ToString();

                            br.BaseStream.Position = 0x218;
                            tbDriveSize.Text = ClassMbr.ConvertBytes(BitConverter.ToUInt64(br.ReadBytes(0x8), 0)).ToString();

                            br.BaseStream.Position = 0x220;
                            tbContentId.Text = (new Guid(br.ReadBytes(0x10))).ToString();

                            br.BaseStream.Position = 0x39C;
                            tbProductID.Text = (new Guid(br.ReadBytes(0x10))).ToString();

                            br.BaseStream.Position = 0x3AC;
                            tbBuildID.Text = (new Guid(br.ReadBytes(0x10))).ToString();

                            br.BaseStream.Position = 0x3BC;
                            ushort PackageVersion1 = BitConverter.ToUInt16(br.ReadBytes(0x2), 0);
                            br.BaseStream.Position = 0x3BE;
                            ushort PackageVersion2 = BitConverter.ToUInt16(br.ReadBytes(0x2), 0);
                            br.BaseStream.Position = 0x3C0;
                            ushort PackageVersion3 = BitConverter.ToUInt16(br.ReadBytes(0x2), 0);
                            br.BaseStream.Position = 0x3C2;
                            ushort PackageVersion4 = BitConverter.ToUInt16(br.ReadBytes(0x2), 0);
                            tbPackageVersion.Text = PackageVersion4 + "." + PackageVersion3 + "." + PackageVersion2 + "." + PackageVersion1;
                            linkCopyContentID.Enabled = true;
                            if (!string.IsNullOrEmpty(tbFilePath.Text))
                            {
                                string filename = Path.GetFileName(tbFilePath.Text).ToLowerInvariant();
                                if (filename != tbContentId.Text.ToLowerInvariant() && !filename.EndsWith(".msixvc")) linkRename.Enabled = true;
                            }
                        }
                        else
                        {
                            MessageBox.Show("不是有效文件", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        br.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("不是有效文件", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LinkCopyContentID_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string sContentID = tbContentId.Text;
            if (!string.IsNullOrEmpty(sContentID))
            {
                Clipboard.SetDataObject(sContentID.ToUpper());
            }
        }

        private void LinkRename_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show(string.Format("是否确认重命名本地文件？\n\n修改前文件名：{0}\n修改后文件名：{1}", Path.GetFileName(tbFilePath.Text), tbContentId.Text.ToUpperInvariant()), "重命名本地文件", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                FileInfo fi = new FileInfo(tbFilePath.Text);
                try
                {
                    fi.MoveTo(Path.GetDirectoryName(tbFilePath.Text) + "\\" + tbContentId.Text.ToUpperInvariant());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("重命名本地文件失败，错误信息：" + ex.Message, "重命名本地文件", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                linkRename.Enabled = false;
            }
        }
        #endregion

        #region 选项卡-游戏
        private void ButGame_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbGameUrl.Text)) return;
            Match result = Regex.Match(tbGameUrl.Text.Trim(), @"/(?<productId>[a-zA-Z0-9]{12})$|/(?<productId>[a-zA-Z0-9]{12})(\?|#)|^(?<productId>[a-zA-Z0-9]{12})$");
            if (result.Success)
            {
                pbGame.Image = pbGame.InitialImage;
                tbGameTitle.Clear();
                tbGameDeveloperName.Clear();
                tbGameCategory.Clear();
                tbGameOriginalReleaseDate.Clear();
                cbGameBundled.Items.Clear();
                tbGamePrice.Clear();
                tbGameXgp1.Clear();
                tbGameLanguages.Clear();
                lvGame.Items.Clear();
                butGame.Enabled = false;
                linkCompare.Enabled = false;
                linkGameWebsite.Enabled = false;
                this.cbGameBundled.SelectedIndexChanged -= new EventHandler(this.CbGameBundled_SelectedIndexChanged);
                Market market = (Market)cbGameMarket.SelectedItem;
                string productId = result.Groups["productId"].Value.ToUpperInvariant();
                ThreadPool.QueueUserWorkItem(delegate { GameSniffer(market, productId); });
            }
            else
            {
                MessageBox.Show("无效 URL/ProductId", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        string query = string.Empty;
        private void TbGameSearch_TextChanged(object sender, EventArgs e)
        {
            string query = tbGameSearch.Text.Trim();
            if (string.IsNullOrEmpty(query))
            {
                lbGameSearch.Items.Clear();
                lbGameSearch.Visible = false;
                this.query = query;
                return;
            }
            if (this.query == query) return;
            this.query = query;
            ThreadPool.QueueUserWorkItem(delegate { GameSearch(query); });
        }

        private void TbGameSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == (int)Keys.Down)
            {
                if (lbGameSearch.Items.Count >= 1)
                {
                    lbGameSearch.Focus();
                    lbGameSearch.SelectedIndex = lbGameSearch.SelectedIndex < lbGameSearch.Items.Count - 1 ? lbGameSearch.SelectedIndex + 1 : lbGameSearch.Items.Count - 1;
                }
            }
            else if (e.KeyValue == (int)Keys.Up)
            {
                if (lbGameSearch.Items.Count >= 1)
                {
                    lbGameSearch.Focus();
                    lbGameSearch.SelectedIndex = lbGameSearch.SelectedIndex > 1 ? lbGameSearch.SelectedIndex - 1 : 0;
                }
            }
        }

        private void TbGameSearch_Leave(object sender, EventArgs e)
        {
            if (lbGameSearch.Focused == false)
            {
                lbGameSearch.Visible = false;
            }
        }

        private void TbGameSearch_Enter(object sender, EventArgs e)
        {
            if (lbGameSearch.Items.Count >= 1)
            {
                lbGameSearch.Visible = true;
            }
        }

        private void LbGameSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == (int)Keys.Enter)
            {
                Product product = (Product)lbGameSearch.SelectedItem;
                lbGameSearch.Visible = false;
                tbGameUrl.Text = "https://www.microsoft.com/store/productId/" + product.id;
                if (butGame.Enabled) ButGame_Click(null, null);
            }
        }

        private void LbGameSearch_DoubleClick(object sender, EventArgs e)
        {
            if (lbGameSearch.SelectedItem != null)
            {
                Product product = (Product)lbGameSearch.SelectedItem;
                lbGameSearch.Visible = false;
                tbGameUrl.Text = "https://www.microsoft.com/store/productId/" + product.id;
                if (butGame.Enabled) ButGame_Click(null, null);
            }
        }

        private void LbGameSearch_Leave(object sender, EventArgs e)
        {
            if (tbGameSearch.Focused == false)
            {
                lbGameSearch.Visible = false;
            }
        }

        private void GameSearch(string query)
        {
            Thread.Sleep(300);
            if (this.query != query) return;
            string language;
            switch (Thread.CurrentThread.CurrentCulture.Name)
            {
                case "zh-HK":
                case "zh-TW":
                    language = "zh-TW";
                    break;
                default:
                    language = "zh-SG";
                    break;
            }
            string url = "https://www.microsoft.com/services/api/v3/suggest?market=" + language + "&clientId=7F27B536-CF6B-4C65-8638-A0F8CBDFCA65&sources=Microsoft-Terms%2CIris-Products%2CDCatAll-Products&filter=ExcludeDCatProducts%3ADCatDevices-Products%2CDCatSoftware-Products%2CDCatBundles-Products%2BClientType%3AStoreWeb&counts=5%2C1%2C5&query=" + ClassWeb.UrlEncode(query);
            SocketPackage socketPackage = ClassWeb.HttpRequest(url, "GET", null, null, true, false, true, null, null, null, ClassWeb.useragent, null, null, null, null, 0, null);
            if (this.query != query) return;
            List<Product> lsProduct = new List<Product>();
            if (Regex.IsMatch(socketPackage.Html, @"^{.+}$", RegexOptions.Singleline))
            {
                JavaScriptSerializer js = new JavaScriptSerializer();
                var json = js.Deserialize<ClassGame.Search>(socketPackage.Html);
                if (json != null && json.ResultSets != null && json.ResultSets.Count >= 1)
                {
                    foreach (var resultSets in json.ResultSets)
                    {
                        foreach (var suggest in resultSets.Suggests)
                        {
                            if (suggest.Source != "Games") continue;
                            var BigCatalogId = Array.FindAll(suggest.Metas.ToArray(), a => a.Key == "BigCatalogId");
                            if (BigCatalogId.Length == 1)
                            {
                                lsProduct.Add(new Product(suggest.Title, BigCatalogId[0].Value));
                            }
                        }
                    }
                }
            }
            this.Invoke(new Action(() =>
            {
                lbGameSearch.Items.Clear();
                if (lsProduct.Count >= 1)
                {
                    int height = (int)(15 * Form1.dpixRatio);
                    lbGameSearch.Items.AddRange(lsProduct.ToArray());
                    lbGameSearch.Height = (lsProduct.Count <= 8 ? lsProduct.Count * height : 8 * height);
                    lbGameSearch.Visible = true;
                }
                else
                {
                    lbGameSearch.Visible = false;
                }
            }));
        }

        private void GameXGPRecentlyAdded(int sort)
        {
            ComboBox cb;
            string siglId1 = string.Empty, siglId2 = string.Empty, text1 = string.Empty, text2 = string.Empty;
            if (sort == 1)
            {
                cb = cbGameXGP1;
                siglId1 = "eab7757c-ff70-45af-bfa6-79d3cfb2bf81";
                siglId2 = "a884932a-f02b-40c8-a903-a008c23b1df1";
                text1 = "最受欢迎 Xbox Game Pass 主机游戏 ({0})";
                text2 = "最受欢迎 Xbox Game Pass 电脑游戏 ({0})";
            }
            else
            {
                cb = cbGameXGP2;
                siglId1 = "f13cf6b4-57e6-4459-89df-6aec18cf0538";
                siglId2 = "163cdff5-442e-4957-97f5-1050a3546511";
                text1 = "近期新增 Xbox Game Pass 主机游戏 ({0})";
                text2 = "近期新增 Xbox Game Pass 电脑游戏 ({0})";
            }
            List<Product> lsProduct1 = new List<Product>();
            List<Product> lsProduct2 = new List<Product>();
            Task[] tasks = new Task[2];
            tasks[0] = new Task(() =>
            {
                lsProduct1 = GetXGPGames(siglId1, text1);
            });
            tasks[1] = new Task(() =>
            {
                lsProduct2 = GetXGPGames(siglId2, text2);
            });
            Array.ForEach(tasks, x => x.Start());
            Task.WaitAll(tasks);
            List<Product> lsProduct = lsProduct1.Union(lsProduct2).ToList<Product>();
            if (lsProduct.Count >= 1)
            {
                this.Invoke(new Action(() =>
                {
                    cb.Items.Clear();
                    cb.Items.AddRange(lsProduct.ToArray());
                    cb.SelectedIndex = 0;
                }));
            }
        }

        private List<Product> GetXGPGames(string siglId, string text)
        {
            List<Product> lsProduct = new List<Product>();
            List<string> lsBundledId = new List<string>();
            string url = "https://catalog.gamepass.com/sigls/v2?id=" + siglId + "&language=zh-Hans&market=US";
            SocketPackage socketPackage = ClassWeb.HttpRequest(url, "GET", null, null, true, false, true, null, null, null, ClassWeb.useragent, null, null, null, null, 0, null, 60000, 60000);
            Match result = Regex.Match(socketPackage.Html, @"\{""id"":""(?<ProductId>[a-zA-Z0-9]{12})""\}");
            while (result.Success)
            {
                lsBundledId.Add(result.Groups["ProductId"].Value.ToLowerInvariant());
                result = result.NextMatch();
            }
            if (lsBundledId.Count >= 1)
            {
                url = "https://displaycatalog.mp.microsoft.com/v7.0/products?bigIds=" + string.Join(",", lsBundledId.ToArray()) + "&market=US&languages=zh-Hans&MS-CV=DGU1mcuYo0WMMp+F.1";
                socketPackage = ClassWeb.HttpRequest(url, "GET", null, null, true, false, true, null, null, null, ClassWeb.useragent, null, null, null, null, 0, null, 60000, 60000);
                if (Regex.IsMatch(socketPackage.Html, @"^{.+}$", RegexOptions.Singleline))
                {
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    var json = js.Deserialize<ClassGame.Game>(socketPackage.Html);
                    if (json != null && json.Products != null && json.Products.Count >= 1)
                    {
                        lsProduct.Add(new Product(string.Format(text, json.Products.Count), "0"));
                        foreach (var product in json.Products)
                        {
                            lsProduct.Add(new Product("  " + product.LocalizedProperties[0].ProductTitle, product.ProductId));
                        }
                    }
                }
            }
            if (lsProduct.Count == 0)
                lsProduct.Add(new Product(string.Format(text, "加载失败"), "0"));
            return lsProduct;
        }

        private void CbGame_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            if (cb.SelectedIndex <= 0) return;
            Product product = (Product)cb.SelectedItem;
            if (product.id == "0") return;
            tbGameUrl.Text = "https://www.microsoft.com/store/productId/" + product.id;
            if (butGame.Enabled) ButGame_Click(null, null);
        }

        private void LinkGameChinese_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FormChinese dialog = new FormChinese();
            dialog.ShowDialog();
            dialog.Dispose();
            if (!string.IsNullOrEmpty(dialog.productid))
            {
                tbGameUrl.Text = "https://www.microsoft.com/store/productId/" + dialog.productid;
                foreach (var item in cbGameMarket.Items)
                {
                    Market market = (Market)item;
                    if (market.lang == "zh-CN")
                    {
                        cbGameMarket.SelectedItem = item;
                        break;
                    }
                }
                if (butGame.Enabled) ButGame_Click(null, null);
            }
        }

        private void GameWithGold()
        {
            ConcurrentDictionary<String, string[]> dicGamesWithGold = new ConcurrentDictionary<String, string[]>();
            SocketPackage socketPackage = ClassWeb.HttpRequest("https://www.xbox.com/en-US/live/gold/js/gwg-globalContent.js", "GET", null, null, true, false, true, null, null, null, ClassWeb.useragent, null, null, null, null, 0, null, 60000, 60000);
            Match result = Regex.Match(Regex.Replace(socketPackage.Html, @"globalContentOld.+", "", RegexOptions.Singleline), @"""(?<langue>[^""]+)"": \{\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyCopytitlenowgame1"": ""(?<keyCopytitlenowgame1>[^""]+)"",\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyLinknowgame1"": ""(?<keyLinknowgame1>[^""]+)"",\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyCopydatesnowgame1"": ""(?<keyCopydatesnowgame1>[^""]+)"",\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyCopytitlenowgame2"": ""(?<keyCopytitlenowgame2>[^""]+)"",\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyLinknowgame2"": ""(?<keyLinknowgame2>[^""]+)"",\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyCopydatesnowgame2"": ""(?<keyCopydatesnowgame2>[^""]+)"",\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyCopytitlenowgame3"": ""(?<keyCopytitlenowgame3>[^""]+)"",\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyLinknowgame3"": ""(?<keyLinknowgame3>[^""]+)"",\n(\s+""[^""]+"": ""[^""]*"",\n)+\s+""keyCopydatesnowgame3"": ""(?<keyCopydatesnowgame3>[^""]+)""");
            while (result.Success)
            {
                string lengue = result.Groups["langue"].Value.ToLowerInvariant();
                string keyCopytitlenowgame1 = result.Groups["keyCopytitlenowgame1"].Value;
                string keyCopytitlenowgame2 = result.Groups["keyCopytitlenowgame2"].Value;
                string keyCopytitlenowgame3 = result.Groups["keyCopytitlenowgame3"].Value;
                string keyLinknowgame1 = result.Groups["keyLinknowgame1"].Value;
                string keyLinknowgame2 = result.Groups["keyLinknowgame2"].Value;
                string keyLinknowgame3 = result.Groups["keyLinknowgame3"].Value;
                string keyCopydatesnowgame1 = result.Groups["keyCopydatesnowgame1"].Value;
                string keyCopydatesnowgame2 = result.Groups["keyCopydatesnowgame2"].Value;
                string keyCopydatesnowgame3 = result.Groups["keyCopydatesnowgame3"].Value;
                if (lengue == "en-sg")
                {
                    string[] detail1 = new string[] { "zh-sg", keyCopytitlenowgame1, keyLinknowgame1, keyCopydatesnowgame1 };
                    string[] detail2 = new string[] { "zh-sg", keyCopytitlenowgame2, keyLinknowgame2, keyCopydatesnowgame2 };
                    string[] detail3 = new string[] { "zh-sg", keyCopytitlenowgame3, keyLinknowgame3, keyCopydatesnowgame3 };
                    dicGamesWithGold.AddOrUpdate(keyLinknowgame1, detail1, (oldkey, oldvalue) => detail1);
                    dicGamesWithGold.AddOrUpdate(keyLinknowgame2, detail2, (oldkey, oldvalue) => detail2);
                    dicGamesWithGold.AddOrUpdate(keyLinknowgame3, detail3, (oldkey, oldvalue) => detail3);
                }
                else
                {
                    if (!dicGamesWithGold.ContainsKey(keyLinknowgame1))
                    {
                        string[] detail1 = new string[] { lengue, keyCopytitlenowgame1, keyLinknowgame1, keyCopydatesnowgame1 };
                        dicGamesWithGold.TryAdd(keyLinknowgame1, detail1);
                    }
                    if (!dicGamesWithGold.ContainsKey(keyLinknowgame2))
                    {
                        string[] detail2 = new string[] { lengue, keyCopytitlenowgame2, keyLinknowgame2, keyCopydatesnowgame2 };
                        dicGamesWithGold.TryAdd(keyLinknowgame2, detail2);
                    }
                    if (!dicGamesWithGold.ContainsKey(keyLinknowgame3))
                    {
                        string[] detail3 = new string[] { lengue, keyCopytitlenowgame3, keyLinknowgame3, keyCopydatesnowgame3 };
                        dicGamesWithGold.TryAdd(keyLinknowgame3, detail3);
                    }
                }
                result = result.NextMatch();
            }
            if (dicGamesWithGold.Count >= 1)
            {
                this.Invoke(new Action(() =>
                {
                    flpGameWithGold.Controls.Clear();
                    foreach (var item in dicGamesWithGold)
                    {
                        LinkLabel lb = new LinkLabel()
                        {
                            Tag = item.Value[0],
                            Text = item.Value[1] + "\n" + item.Value[3].Replace(" ", ""),
                            TextAlign = ContentAlignment.TopCenter,
                            AutoSize = true,
                            Parent = this.flpGameWithGold
                        };
                        lb.Links.Add(0, item.Value[1].Length, Regex.Replace(item.Value[2], @"/p/", "/" + item.Value[0] + "/p/"));
                        lb.LinkClicked += new LinkLabelLinkClickedEventHandler(this.LinkGameWithGold_LinkClicked);
                    }
                }));
            }
        }

        private void LinkGameWithGold_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel lb = sender as LinkLabel;
            string langue = lb.Tag.ToString();
            tbGameUrl.Text = e.Link.LinkData as string;
            bool find = false;
            foreach (var item in cbGameMarket.Items)
            {
                Market market = (Market)item;
                if (market.lang.ToLowerInvariant() == langue)
                {
                    cbGameMarket.SelectedItem = item;
                    find = true;
                    break;
                }
            }
            if (!find)
            {
                cbGameMarket.Items.Add(new Market(langue, Regex.Replace(langue, "^[^-]+-", ""), langue));
                cbGameMarket.SelectedIndex = cbGameMarket.Items.Count - 1;
            }
            if (butGame.Enabled) ButGame_Click(null, null);
        }

        private void CbGameBundled_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbGameBundled.SelectedIndex < 0) return;
            tbGameTitle.Clear();
            tbGameDeveloperName.Clear();
            tbGameCategory.Clear();
            tbGameOriginalReleaseDate.Clear();
            tbGamePrice.Clear();
            tbGameXgp1.Clear();
            tbGameLanguages.Clear();
            lvGame.Items.Clear();
            linkCompare.Enabled = false;
            linkGameWebsite.Enabled = false;

            var market = (Market)cbGameBundled.Tag;
            var json = (ClassGame.Game)gbGameInfo.Tag;
            GameAnalyzer(market, json, cbGameBundled.SelectedIndex);
        }
        
        internal static ConcurrentDictionary<String, Double> dicExchangeRate = new ConcurrentDictionary<String, Double>();
        private void GameSniffer(Market market, string productId)
        {
            cbGameBundled.Tag = market;
            string url = "https://displaycatalog.mp.microsoft.com/v7.0/products?bigIds=" + productId + "&market=" + market.code + "&languages=" + market.lang + ",neutral&MS-CV=DGU1mcuYo0WMMp+F.1";
            SocketPackage socketPackage = ClassWeb.HttpRequest(url, "GET", null, null, true, false, true, null, null, null, ClassWeb.useragent, null, null, null, null, 0, null);
            if (Regex.IsMatch(socketPackage.Html, @"^{.+}$", RegexOptions.Singleline))
            {
                JavaScriptSerializer js = new JavaScriptSerializer();
                var json = js.Deserialize<ClassGame.Game>(socketPackage.Html);
                if (json != null && json.Products != null && json.Products.Count >= 1)
                {
                    GameAnalyzer(market, json, 0);
                }
                else
                {
                    this.Invoke(new Action(() =>
                    {
                        MessageBox.Show("无效 URL/ProductId", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        butGame.Enabled = true;
                    }));
                }
            }
            else
            {
                this.Invoke(new Action(() =>
                {
                    MessageBox.Show("无法连接服务器，请稍候再试。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    butGame.Enabled = true;
                }));
            }
        }

        private void GameAnalyzer(Market market, ClassGame.Game json, int index)
        {
            string title = string.Empty, developerName = string.Empty, description = string.Empty;
            var product = json.Products[index];
            List<string> bundledId = new List<string>();
            List<ListViewItem> lsDownloadUrl = new List<ListViewItem>();
            var localizedPropertie = product.LocalizedProperties;
            if (localizedPropertie != null && localizedPropertie.Count >= 1)
            {
                title = localizedPropertie[0].ProductTitle;
                developerName = localizedPropertie[0].DeveloperName;
                description = localizedPropertie[0].ProductDescription;
                string imageUri = string.Empty, tmpUri = null;
                int imgMin = 0;
                foreach (var image in localizedPropertie[0].Images)
                {
                    if (image.ImagePurpose == "Logo" || image.ImagePurpose == "BoxArt") //Poster, BrandedKeyArt
                    {
                        if (image.Width >= 300 && image.Width == image.Height)
                        {
                            if (string.IsNullOrEmpty(imageUri))
                            {
                                imgMin = image.Width;
                                imageUri = image.Uri;
                            }
                            else if (image.Width < imgMin)
                            {
                                imgMin = image.Width;
                                imageUri = image.Uri;
                            }
                        }
                    }
                    if (image.Width >= 300 && image.Width == image.Height)
                        tmpUri = image.Uri;
                }
                if (string.IsNullOrEmpty(imageUri)) imageUri = tmpUri;
                if (!string.IsNullOrEmpty(imageUri))
                {
                    try
                    {
                        pbGame.LoadAsync("http:" + imageUri);
                    }
                    catch { }
                }
            }

            string originalReleaseDate = string.Empty;
            var marketProperties = product.MarketProperties;
            if (marketProperties != null && marketProperties.Count >= 1)
            {
                originalReleaseDate = marketProperties[0].OriginalReleaseDate.ToString("d");
            }

            string category = string.Empty;
            var properties = product.Properties;
            if (properties != null )
            {
                category = properties.Category;
            }

            string languages = string.Empty;
            if (product.DisplaySkuAvailabilities != null)
            {
                foreach (var displaySkuAvailabilitie in product.DisplaySkuAvailabilities)
                {
                    if (displaySkuAvailabilitie.Sku.SkuType == "full")
                    {
                        if (displaySkuAvailabilitie.Sku.Properties.Packages != null)
                        {
                            foreach (var Packages in displaySkuAvailabilitie.Sku.Properties.Packages)
                            {
                                List<ClassGame.PlatformDependencies> platformDependencie = Packages.PlatformDependencies;
                                List<ClassGame.PackageDownloadUris> packageDownloadUri = Packages.PackageDownloadUris;
                                if (platformDependencie != null && packageDownloadUri != null && Packages.PlatformDependencies.Count >= 1 && packageDownloadUri.Count >= 1)
                                {
                                    string url = packageDownloadUri[0].Uri;
                                    if (url == "https://productingestionbin1.blob.core.windows.net") url = "";
                                    switch (platformDependencie[0].PlatformName)
                                    {
                                        case "Windows.Xbox":
                                            if (Packages.PackageRank == 51000) //packageDownloadUri[0].Uri.EndsWith("_xs.xvc")
                                                lsDownloadUrl.Add(new ListViewItem(new string[] { "Xbox Series X|S", market.name, ClassMbr.ConvertBytes(Packages.MaxDownloadSizeInBytes), url }));
                                            else
                                                lsDownloadUrl.Add(new ListViewItem(new string[] { "Xbox One", market.name, ClassMbr.ConvertBytes(Packages.MaxDownloadSizeInBytes), url }));
                                            break;
                                        case "Windows.Desktop":
                                            lsDownloadUrl.Add(new ListViewItem(new string[] { "微软商店(PC)", market.name, ClassMbr.ConvertBytes(Packages.MaxDownloadSizeInBytes), url }));
                                            break;
                                    }
                                    if (Packages.Languages != null) languages = string.Join(", ", Packages.Languages);
                                }
                            }
                        }
                        if (displaySkuAvailabilitie.Sku.Properties.BundledSkus != null)
                        {
                            foreach (var BundledSkus in displaySkuAvailabilitie.Sku.Properties.BundledSkus)
                            {
                                bundledId.Add(BundledSkus.BigId);
                            }
                        }
                        break;
                    }
                }
            }

            List<Product> lsProduct = new List<Product>();
            if (bundledId.Count >= 1 && json.Products.Count == 1)
            {
                string url = "https://displaycatalog.mp.microsoft.com/v7.0/products?bigIds=" + string.Join(",", bundledId.ToArray()) + "&market=" + market.code + "&languages=" + market.lang + ",neutral&MS-CV=DGU1mcuYo0WMMp+F.1";
                SocketPackage socketPackage = ClassWeb.HttpRequest(url, "GET", null, null, true, false, true, null, null, null, ClassWeb.useragent, null, null, null, null, 0, null);
                if (Regex.IsMatch(socketPackage.Html, @"^{.+}$", RegexOptions.Singleline))
                {
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    var json2 = js.Deserialize<ClassGame.Game>(socketPackage.Html);
                    if (json2 != null && json2.Products != null && json2.Products.Count >= 1)
                    {
                        json.Products.AddRange(json2.Products);
                        lsProduct.Add(new Product("在此捆绑包中(" + json2.Products.Count + ")", ""));
                        foreach (var product2 in json2.Products)
                        {
                            lsProduct.Add(new Product(product2.LocalizedProperties[0].ProductTitle, product2.ProductId));
                        }
                    }
                }
            }

            if (index == 0) gbGameInfo.Tag = json;
            string CurrencyCode = product.DisplaySkuAvailabilities[0].Availabilities[0].OrderManagementData.Price.CurrencyCode.ToUpperInvariant();
            double MSRP = product.DisplaySkuAvailabilities[0].Availabilities[0].OrderManagementData.Price.MSRP;
            double ListPrice_1 = product.DisplaySkuAvailabilities[0].Availabilities[0].OrderManagementData.Price.ListPrice;
            double ListPrice_2 = product.DisplaySkuAvailabilities[0].Availabilities.Count >= 2 ? product.DisplaySkuAvailabilities[0].Availabilities[1].OrderManagementData.Price.ListPrice : 0;
            double WholesalePrice_1 = product.DisplaySkuAvailabilities[0].Availabilities[0].OrderManagementData.Price.WholesalePrice;
            double WholesalePrice_2 = product.DisplaySkuAvailabilities[0].Availabilities.Count >= 2 ? product.DisplaySkuAvailabilities[0].Availabilities[1].OrderManagementData.Price.WholesalePrice : 0;
            if (ListPrice_1 > MSRP) MSRP = ListPrice_1;
            if (!string.IsNullOrEmpty(CurrencyCode) && MSRP > 0 && CurrencyCode != "CNY" && !dicExchangeRate.ContainsKey(CurrencyCode))
            {
                ClassGame.ExchangeRate(CurrencyCode);
            }
            double ExchangeRate = dicExchangeRate.ContainsKey(CurrencyCode) ? dicExchangeRate[CurrencyCode] : 0;

            this.Invoke(new Action(() =>
            {
                tbGameTitle.Text = title;
                tbGameDeveloperName.Text = developerName;
                tbGameCategory.Text = category;
                tbGameOriginalReleaseDate.Text = originalReleaseDate;
                if (lsProduct.Count >= 1)
                {
                    cbGameBundled.Items.AddRange(lsProduct.ToArray());
                    cbGameBundled.SelectedIndex = 0;
                    this.cbGameBundled.SelectedIndexChanged += new EventHandler(this.CbGameBundled_SelectedIndexChanged);
                }
                tbGameXgp1.Text = description;
                tbGameLanguages.Text = languages;
                if (MSRP > 0)
                {
                    
                    StringBuilder sb = new StringBuilder();
                    sb.Append(string.Format("币种: {0}, 建议零售价: {1}", CurrencyCode, String.Format("{0:N}", MSRP)));
                    if (ExchangeRate > 0)
                    {
                        sb.Append(string.Format("({0})", String.Format("{0:N}", MSRP * ExchangeRate)));
                    }
                    if (ListPrice_1 > 0 && ListPrice_1 != MSRP)
                    {
                        sb.Append(string.Format(", 普通折扣{0}%: {1}", Math.Round(ListPrice_1 / MSRP * 100, 0, MidpointRounding.AwayFromZero), String.Format("{0:N}", ListPrice_1)));
                        if (ExchangeRate > 0)
                        {
                            sb.Append(string.Format("({0})", String.Format("{0:N}", ListPrice_1 * ExchangeRate)));
                        }
                    }
                    if (ListPrice_2 > 0 && ListPrice_2 < ListPrice_1 && ListPrice_2 != MSRP)
                    {
                        string member = (product.DisplaySkuAvailabilities[0].Availabilities[1].Properties.MerchandisingTags != null && product.DisplaySkuAvailabilities[0].Availabilities[1].Properties.MerchandisingTags[0] == "LegacyDiscountEAAccess") ? "EA Play" : "金会员";
                        sb.Append(string.Format(", {0}折扣{1}%: {2}", member, Math.Round(ListPrice_2 / MSRP * 100, 0, MidpointRounding.AwayFromZero), String.Format("{0:N}", ListPrice_2)));
                        if (ExchangeRate > 0)
                        {
                            sb.Append(string.Format("({0})", String.Format("{0:N}", ListPrice_2 * ExchangeRate)));
                        }
                    }
                    if (WholesalePrice_1 > 0)
                    {
                        sb.Append(string.Format(", 批发价: {0}", String.Format("{0:N}", WholesalePrice_1)));
                        if (ExchangeRate > 0)
                        {
                            sb.Append(string.Format("({0})", String.Format("{0:N}", WholesalePrice_1 * ExchangeRate)));
                        }
                        if (WholesalePrice_2 > 0 && WholesalePrice_2 < WholesalePrice_1)
                        {
                            sb.Append(string.Format(", 批发价折扣{0}%: {1}", Math.Round(WholesalePrice_2 / WholesalePrice_1 * 100, 0, MidpointRounding.AwayFromZero), String.Format("{0:N}", WholesalePrice_2)));
                            if (ExchangeRate > 0)
                            {
                                sb.Append(string.Format("({0})", String.Format("{0:N}", WholesalePrice_2 * ExchangeRate)));
                            }
                        }
                    }
                    if (ExchangeRate > 0)
                    {
                        sb.Append(string.Format(", CNY汇率: {0}", ExchangeRate));
                    }
                    tbGamePrice.Text = sb.ToString();
                    linkCompare.Enabled = true;
                }
                if (lsDownloadUrl.Count >= 1)
                {
                    lsDownloadUrl.Sort((x, y) => string.Compare(x.SubItems[0].Text, y.SubItems[0].Text));
                    lvGame.Items.AddRange(lsDownloadUrl.ToArray());
                }
                butGame.Enabled = true;
                linkGameWebsite.Enabled = true;
            }));
        }

        private void LinkGameWebsite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var market = (Market)cbGameBundled.Tag;
            var json = (ClassGame.Game)gbGameInfo.Tag;
            int index = cbGameBundled.SelectedIndex == -1 ? 0 : cbGameBundled.SelectedIndex;
            Process.Start("https://www.microsoft.com/" + market.lang + "/p/_/" + json.Products[index].ProductId);
        }

        private void LinkCompare_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            int index = cbGameBundled.SelectedIndex == -1 ? 0 : cbGameBundled.SelectedIndex;
            FormCompare dialog = new FormCompare(gbGameInfo.Tag, index);
            dialog.ShowDialog();
            dialog.Dispose();
        }

        private void LvGame_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (lvGame.SelectedItems.Count == 1 && !string.IsNullOrEmpty(lvGame.SelectedItems[0].SubItems[3].Text))
                {
                    tsmCopyUrl2.Enabled = Regex.IsMatch(lvGame.SelectedItems[0].SubItems[3].Text, @"\.xboxlive\.com");
                    contextMenuStrip2.Show(MousePosition.X, MousePosition.Y);
                }
            }
        }

        private void TsmCopyUrl1_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(lvGame.SelectedItems[0].SubItems[3].Text);
        }

        private void TsmCopyUrl2_Click(object sender, EventArgs e)
        {
            string url = lvGame.SelectedItems[0].SubItems[3].Text;
            url = url.Replace(".xboxlive.com", ".xboxlive.cn");
            Clipboard.SetDataObject(url);
        }

        private void LinkAppxAdd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            tabControl1.SelectedIndex = 6;
            tbAppxFilePath.Focus();
        }
        #endregion

        #region 选项卡-工具
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x219)
            {
                switch (m.WParam.ToInt32())
                {
                    case 0x8000: //U盘插入
                    case 0x8004: //U盘卸载
                        LinkRefreshDrive_LinkClicked(null, null);
                        break;
                    default:
                        break;
                }
            }
            base.WndProc(ref m);
        }

        private void LinkRefreshDrive_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            cbDrive.Items.Clear();
            DriveInfo[] driverList = Array.FindAll(DriveInfo.GetDrives(), a => a.DriveType == DriveType.Removable);
            if (driverList.Length >= 1)
            {
                cbDrive.Items.AddRange(driverList);
                cbDrive.SelectedIndex = 0;
            }
            else
            {
                labelStatusDrive.Text = "当前U盘状态：没有找到U盘";
            }
        }

        private void CbDrive_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbDrive.Items.Count >= 1)
            {
                string driverName = cbDrive.Text;
                DriveInfo driveInfo = new DriveInfo(driverName);
                if (driveInfo.DriveType == DriveType.Removable)
                {
                    if (driveInfo.IsReady && driveInfo.DriveFormat == "NTFS")
                    {
                        List<string> listStatus = new List<string>();
                        if (File.Exists(driverName + "$ConsoleGen8Lock"))
                            listStatus.Add(rbXboxOne.Text + " 回国");
                        else if (File.Exists(driverName + "$ConsoleGen8"))
                            listStatus.Add(rbXboxOne.Text + " 出国");
                        if (File.Exists(driverName + "$ConsoleGen9Lock"))
                            listStatus.Add(rbXboxSeries.Text + " 回国");
                        else if (File.Exists(driverName + "$ConsoleGen9"))
                            listStatus.Add(rbXboxSeries.Text + " 出国");
                        if (listStatus.Count >= 1)
                            labelStatusDrive.Text = "当前U盘状态：" + string.Join(", ", listStatus.ToArray());
                        else
                            labelStatusDrive.Text = "当前U盘状态：未转换";
                    }
                    else
                    {
                        labelStatusDrive.Text = "当前U盘状态：不是NTFS格式";
                    }
                }
            }
        }

        private void ButConsoleRegionUnlock_Click(object sender, EventArgs e)
        {
            ConsoleRegion(true);
        }

        private void ButConsoleRegionLock_Click(object sender, EventArgs e)
        {
            ConsoleRegion(false);
        }

        private void ConsoleRegion(bool unlock)
        {
            if (cbDrive.Items.Count == 0)
            {
                MessageBox.Show("请插入U盘。", "没有找到U盘", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            labelStatusDrive.Text = "当前U盘状态：制作中，请稍候...";
            string driverName = cbDrive.Text;
            DriveInfo driveInfo = new DriveInfo(driverName);
            if (driveInfo.DriveType == DriveType.Removable)
            {
                if (!driveInfo.IsReady || driveInfo.DriveFormat != "NTFS")
                {
                    string show, caption, cmd;
                    if (driveInfo.IsReady && driveInfo.DriveFormat == "FAT32")
                    {
                        show = "当前U盘格式 " + driveInfo.DriveFormat + "，是否把U盘转换为 NTFS 格式？\n\n注意，如果U盘有重要数据请先备份!\n\n当前U盘位置： " + driverName + "，容量：" + ClassMbr.ConvertBytes(Convert.ToUInt64(driveInfo.TotalSize)) + "\n取消转换请按\"否(N)\"";
                        caption = "转换U盘格式";
                        cmd = "convert " + Regex.Replace(driverName, @"\\$", "") + " /fs:ntfs /x";
                    }
                    else
                    {
                        show = "当前U盘格式 " + (driveInfo.IsReady ? driveInfo.DriveFormat : "RAW") + "，是否把U盘格式化为 NTFS？\n\n警告，格式化将删除U盘中的所有文件!\n警告，格式化将删除U盘中的所有文件!\n警告，格式化将删除U盘中的所有文件!\n\n当前U盘位置： " + driverName + "，容量：" + (driveInfo.IsReady ? ClassMbr.ConvertBytes(Convert.ToUInt64(driveInfo.TotalSize)) : "未知") + "\n取消格式化请按\"否(N)\"";
                        caption = "格式化U盘";
                        cmd = "format " + Regex.Replace(driverName, @"\\$", "") + " /fs:ntfs /q";
                    }
                    if (MessageBox.Show(show, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        string outputString;
                        using (Process p = new Process())
                        {
                            p.StartInfo.FileName = "cmd.exe";
                            p.StartInfo.UseShellExecute = false;
                            p.StartInfo.RedirectStandardInput = true;
                            p.StartInfo.RedirectStandardError = true;
                            p.StartInfo.RedirectStandardOutput = true;
                            p.StartInfo.CreateNoWindow = true;
                            p.Start();

                            p.StandardInput.WriteLine(cmd);
                            p.StandardInput.WriteLine("exit");

                            p.StandardInput.Close();
                            outputString = p.StandardOutput.ReadToEnd();
                            p.WaitForExit();
                            p.Close();
                        }
                    }
                }
                if (driveInfo.IsReady && driveInfo.DriveFormat == "NTFS")
                {
                    if (File.Exists(driverName + "$ConsoleGen8"))
                        File.Delete(driverName + "$ConsoleGen8");
                    if (File.Exists(driverName + "$ConsoleGen9"))
                        File.Delete(driverName + "$ConsoleGen9");
                    if (File.Exists(driverName + "$ConsoleGen8Lock"))
                        File.Delete(driverName + "$ConsoleGen8Lock");
                    if (File.Exists(driverName + "$ConsoleGen9Lock"))
                        File.Delete(driverName + "$ConsoleGen9Lock");
                    if (rbXboxOne.Checked)
                    {
                        using (File.Create(driverName + (unlock ? "$ConsoleGen8" : "$ConsoleGen8Lock"))) { }
                    }
                    if (rbXboxSeries.Checked)
                    {
                        using (File.Create(driverName + (unlock ? "$ConsoleGen9" : "$ConsoleGen9Lock"))) { }
                    }
                }
                else
                {
                    MessageBox.Show("U盘不是NTFS格式，请重新格式化NTFS格式后再转换。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                CbDrive_SelectedIndexChanged(null, null);
            }
            else
            {
                labelStatusDrive.Text = "当前U盘状态：" + driverName + " 设备不存在";
            }
        }

        private void LinkAppxRefreshDrive_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            cbAppxDrive.Items.Clear();
            DriveInfo[] driverList = Array.FindAll(DriveInfo.GetDrives(), a => a.DriveType != DriveType.Removable);
            if (driverList.Length >= 1)
            {
                cbAppxDrive.Items.AddRange(driverList);
                cbAppxDrive.SelectedIndex = 0;
            }
        }

        private void ButAppxOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Open an Xbox Package"
            };
            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string sFilePath = ofd.FileName;
            tbAppxFilePath.Text = sFilePath;
        }

        private void ButAppxInstall_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbAppxFilePath.Text)) return;
            if (Environment.OSVersion.Version.Major < 10)
            {
                MessageBox.Show("需要Win10操作系统。", "操作系统版本过低", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            using (FileStream fs = File.Create(".install_appx.ps1"))
            {
                Byte[] byteArray = new UTF8Encoding(true).GetBytes("Add-AppxPackage -Path \"" + tbAppxFilePath.Text + "\" -Volume \"" + cbAppxDrive.Text + "\"");
                fs.Write(byteArray, 0, byteArray.Length);
                fs.Close();
            }
            File.SetAttributes(".install_appx.ps1", FileAttributes.Hidden);
            using (Process p = new Process())
            {
                p.StartInfo.FileName = "cmd.exe";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.CreateNoWindow = false;
                p.Start();
                p.StandardInput.WriteLine("powershell -executionpolicy remotesigned -file \".install_appx.ps1\"");
                p.StandardInput.WriteLine("del /a/f/q \".install_appx.ps1\"");
                p.StandardInput.WriteLine("exit");
            }
            tbAppxFilePath.Clear();
        }
        #endregion

        private void Dgv_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, dgv.RowHeadersWidth - 1, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), dgv.RowHeadersDefaultCellStyle.Font, rectangle, dgv.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        delegate void CallbackTextBox(TextBox tb, string str);
        public void SetTextBox(TextBox tb, string str)
        {
            if (tb.InvokeRequired)
            {
                CallbackTextBox d = new CallbackTextBox(SetTextBox);
                Invoke(d, new object[] { tb, str });
            }
            else tb.Text = str;
        }

        delegate void CallbackSaveLog(string status, string content, string ip, int argb);
        public void SaveLog(string status, string content, string ip, int argb = 0)
        {
            if (lvLog.InvokeRequired)
            {
                CallbackSaveLog d = new CallbackSaveLog(SaveLog);
                Invoke(d, new object[] { status, content, ip, argb });
            }
            else
            {
                ListViewItem listViewItem = new ListViewItem(new string[] { status, content, ip, string.Format("{0:T}", DateTime.Now) });
                if (argb >= 1) listViewItem.ForeColor = Color.FromArgb(argb);
                lvLog.Items.Insert(0, listViewItem);
            }
        }

        class ListViewNF : ListView
        {
            public ListViewNF()
            {
                // 开启双缓冲
                this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);

                // Enable the OnNotifyMessage event so we get a chance to filter out 
                // Windows messages before they get to the form's WndProc
                this.SetStyle(ControlStyles.EnableNotifyMessage, true);
            }

            protected override void OnNotifyMessage(Message m)
            {
                //Filter out the WM_ERASEBKGND message
                if (m.Msg != 0x14)
                {
                    base.OnNotifyMessage(m);
                }
            }
        }
    }
}