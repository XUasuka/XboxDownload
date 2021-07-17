using System;
using System.Collections.Generic;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XboxDownload
{
    class DnsListen
    {
        private readonly Form1 parentForm;
        Socket socket = null;

        public DnsListen(Form1 parentForm)
        {
            this.parentForm = parentForm;
        }

        public void Listen()
        {
            IPEndPoint iPEndPoint = null;
            if (string.IsNullOrEmpty(Properties.Settings.Default.DnsIP))
            {
                string priorityIp = Regex.Replace(Properties.Settings.Default.LocalIP, @"\d+$", "");
                bool succeed = false;
                foreach (NetworkInterface adapter in NetworkInterface.GetAllNetworkInterfaces())
                {
                    IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
                    foreach (IPAddress dns in adapterProperties.DnsAddresses)
                    {
                        if (dns.AddressFamily == AddressFamily.InterNetwork)
                        {
                            if (iPEndPoint == null)
                            {
                                iPEndPoint = new IPEndPoint(dns, 53);
                                if (dns.ToString().StartsWith(priorityIp))
                                {
                                    succeed = true;
                                    break;
                                }
                            }
                            else if (dns.ToString().StartsWith(priorityIp))
                            {
                                iPEndPoint = new IPEndPoint(dns, 53);
                                succeed = true;
                                break;
                            }
                        }
                    }
                    if (succeed) break;
                }
                if (iPEndPoint == null)
                    iPEndPoint = new IPEndPoint(IPAddress.Parse("114.114.114.114"), 53);
                parentForm.SetTextBox(parentForm.tbDnsIP, iPEndPoint.Address.ToString());
            }
            else
            {
                iPEndPoint = new IPEndPoint(IPAddress.Parse(Properties.Settings.Default.DnsIP), 53);
            }
            IPEndPoint ipe = new IPEndPoint(Properties.Settings.Default.ListenIP == 0 ? IPAddress.Parse(Properties.Settings.Default.LocalIP) : IPAddress.Any, 53);
            socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
            try
            {
                socket.Bind(ipe);
            }
            catch (SocketException ex)
            {
                parentForm.Invoke(new Action(() =>
                {
                    parentForm.pictureBox1.Image = Properties.Resources.Xbox3;
                    MessageBox.Show(String.Format("启用DNS服务失败!\n错误信息: {0}", ex.Message), "启用DNS服务失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
                return;
            }

            Byte[] comIP = null;
            if (!string.IsNullOrEmpty(Properties.Settings.Default.ComIP))
            {
                string[] ips = Properties.Settings.Default.ComIP.Split('.');
                comIP = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
            }
            else
            {
                parentForm.SetTextBox(parentForm.tbComIP, Properties.Settings.Default.LocalIP);
                string[] ips = Properties.Settings.Default.LocalIP.Split('.');
                comIP = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
            }
            Byte[] cnIP = null;
            if (!string.IsNullOrEmpty(Properties.Settings.Default.CnIP))
            {
                string[] ips = Properties.Settings.Default.CnIP.Split('.');
                cnIP = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
            }
            else
            {
                string ip = ClassWeb.HostToIP("assets1.xboxlive.cn", Properties.Settings.Default.DnsIP);
                if (!string.IsNullOrEmpty(ip))
                {
                    parentForm.SetTextBox(parentForm.tbCnIP, ip);
                    string[] ips = ip.Split('.');
                    cnIP = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
                }
            }
            Byte[] appIP = null;
            if (!string.IsNullOrEmpty(Properties.Settings.Default.AppIP))
            {
                string[] ips = Properties.Settings.Default.AppIP.Split('.');
                appIP = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
            }
            else
            {
                string ip = ClassWeb.HostToIP("tlu.dl.delivery.mp.microsoft.com", Properties.Settings.Default.DnsIP);
                if (!string.IsNullOrEmpty(ip))
                {
                    parentForm.SetTextBox(parentForm.tbAppIP, ip);
                    string[] ips = ip.Split('.');
                    appIP = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
                }
            }
            Byte[] eaIP = null;
            if (!string.IsNullOrEmpty(Properties.Settings.Default.EAIP))
            {
                string[] ips = Properties.Settings.Default.EAIP.Split('.');
                eaIP = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
            }
            else
            {
                string ip = ClassWeb.HostToIP("origin-a.akamaihd.net", Properties.Settings.Default.DnsIP);
                if (!string.IsNullOrEmpty(ip))
                {
                    parentForm.SetTextBox(parentForm.tbEAIP, ip);
                    string[] ips = ip.Split('.');
                    eaIP = new byte[4] { byte.Parse(ips[0]), byte.Parse(ips[1]), byte.Parse(ips[2]), byte.Parse(ips[3]) };
                }
            }
            while (Form1.bServiceFlag)
            {
                try
                {
                    var client = (EndPoint)new IPEndPoint(IPAddress.Any, 0);
                    var buff = new byte[512];
                    int read = socket.ReceiveFrom(buff, ref client);
                    Task.Factory.StartNew(() =>
                    {
                        var dns = new DNS(buff, read);
                        if (dns.QR == 0 && dns.Opcode == 0 && dns.Querys.Count == 1 && (dns.Querys[0].QueryType == QueryType.A || dns.Querys[0].QueryType == QueryType.AAAA))
                        {
                            string queryName = dns.Querys[0].QueryName.ToLower();
                            Byte[] ip = null;
                            int argb = 0;
                            switch (queryName)
                            {
                                case "assets1.xboxlive.com":
                                case "7.assets1.xboxlive.com":
                                case "assets2.xboxlive.com":
                                case "dlassets.xboxlive.com":
                                case "dlassets2.xboxlive.com":
                                case "d1.xboxlive.com":
                                case "d2.xboxlive.com":
                                case "xvcf1.xboxlive.com":
                                case "xvcf2.xboxlive.com":
                                    ip = comIP;
                                    argb = 0x008000;
                                    break;
                                case "assets1.xboxlive.cn":
                                case "assets2.xboxlive.cn":
                                case "dlassets.xboxlive.cn":
                                case "dlassets2.xboxlive.cn":
                                case "d1.xboxlive.cn":
                                case "d2.xboxlive.cn":
                                    ip = cnIP;
                                    argb = 0x008000;
                                    break;
                                case "dl.delivery.mp.microsoft.com":
                                case "tlu.dl.delivery.mp.microsoft.com":
                                    ip = appIP;
                                    argb = 0x008000;
                                    break;
                                case "origin-a.akamaihd.net":
                                    ip = eaIP;
                                    argb = 0x008000;
                                    break;
                                default:
                                    if (Form1.dicDomain.ContainsKey(queryName))
                                        ip = Form1.dicDomain[queryName];
                                    break;
                            }
                            if (Form1.bRecordLog) parentForm.SaveLog("DNS 查询", queryName, ((IPEndPoint)client).Address.ToString(), argb);
                            if (ip != null)
                            {
                                if (dns.Querys[0].QueryType == QueryType.A)
                                {
                                    dns.QR = 1;
                                    dns.RA = 1;
                                    dns.RD = 1;
                                    dns.ResouceRecords = new List<ResouceRecord>
                                    {
                                        new ResouceRecord
                                        {
                                            Datas = ip,
                                            TTL = 100,
                                            QueryClass = 1,
                                            QueryType = QueryType.A
                                        }
                                    };
                                    socket.SendTo(dns.ToBytes(), client);
                                    return;
                                }
                                else // 屏蔽IPv6
                                {
                                    socket.SendTo(new byte[0], client);
                                    return;
                                }
                            }
                        }
                        try
                        {
                            var proxy = new UdpClient();
                            proxy.Client.ReceiveTimeout = 6000;
                            proxy.Connect(iPEndPoint);
                            proxy.Send(buff, read);
                            var bytes = proxy.Receive(ref iPEndPoint);
                            socket.SendTo(bytes, client);
                        }
                        catch (Exception ex)
                        {
                            if (Form1.bRecordLog) parentForm.SaveLog("DNS 查询", ex.Message, ((IPEndPoint)client).Address.ToString());
                        }
                    });
                }
                catch { }
            }
        }

        public void Close()
        {
            if (socket != null)
            {
                socket.Close();
                socket.Dispose();
                socket = null;
            }
        }
    }

    public enum QueryType
    {
        A = 1,
        NS = 2,
        MD = 3,
        MF = 4,
        CNAME = 5,
        SOA = 6,
        MB = 7,
        MG = 8,
        MR = 9,
        WKS = 11,
        PTR = 12,
        HINFO = 13,
        MINFO = 14,
        MX = 15,
        TXT = 16,
        AAAA = 28,
        AXFR = 252,
        ANY = 255
    }

    public class Query
    {
        public string QueryName { get; set; }
        public QueryType QueryType { get; set; }
        public Int16 QueryClass { get; set; }

        public Query()
        {
        }

        public Query(Func<int, byte[]> read)
        {
            var name = new StringBuilder();
            var length = read(1)[0];
            while (length != 0)
            {
                for (var i = 0; i < length; i++)
                {
                    name.Append((char)read(1)[0]);
                }
                length = read(1)[0];
                if (length != 0)
                    name.Append(".");
            }
            QueryName = name.ToString();

            QueryType = (QueryType)IPAddress.NetworkToHostOrder(BitConverter.ToInt16(read(2), 0));
            QueryClass = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(read(2), 0));
        }

        public virtual byte[] ToBytes()
        {
            var list = new List<byte>();

            var a = QueryName.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            for (var i = 0; i < a.Length; i++)
            {
                list.Add((byte)a[i].Length);
                for (var j = 0; j < a[i].Length; j++)
                    list.Add((byte)a[i][j]);
            }
            list.Add(0);

            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((Int16)QueryType)));
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder(QueryClass)));

            return list.ToArray();
        }
    }

    public class ResouceRecord : Query
    {
        public Int16 Point { get; set; }
        public Int32 TTL { get; set; }
        public byte[] Datas { get; set; }

        public ResouceRecord() : base()
        {
            var bytes = new byte[] { 0xc0, 0x0c };
            Point = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(bytes, 0));
        }

        public ResouceRecord(Func<int, byte[]> read) : base()
        {

            TTL = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(read(4), 0));
            var length = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(read(2), 0));
            Datas = read(length);

        }
        public override byte[] ToBytes()
        {
            var list = new List<byte>();
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder(Point)));
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((Int16)QueryType)));
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder(QueryClass)));
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder(TTL)));
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((Int16)Datas.Length)));
            list.AddRange(Datas);

            return list.ToArray();
        }
    }

    public class DNS
    {
        public Int16 标志 { get; set; }
        public int QR { get; set; }     //0表示查询报文 1表示响应报文
        public int Opcode { get; set; } //0表示标准查询,1表示反向查询,2表示服务器状态请求
        public int AA { get; set; }  //授权回答
        public int TC { get; set; } //表示可截断的
        public int RD { get; set; } //表示期望递归 
        public int RA { get; set; } //表示可用递归
        public int Rcode { get; set; } //0表示没有错误,3表示名字错误

        public List<Query> Querys { get; set; }  //问题数
        public List<ResouceRecord> ResouceRecords { get; set; }  //资源记录数
        public Int16 授权资源记录数 { get; set; }
        public Int16 额外资源记录数 { get; set; }

        public byte[] ToBytes()
        {
            var list = new List<byte>();
            var bytes = BitConverter.GetBytes(IPAddress.HostToNetworkOrder(标志));
            list.AddRange(bytes);
            var b = new byte();
            b = b.SetBits(QR, 0, 1)
                .SetBits(Opcode, 1, 4)
                .SetBits(AA, 5, 1)
                .SetBits(TC, 6, 1);

            b = b.SetBits(RD, 7, 1);
            list.Add(b);
            b = new byte();
            b = b.SetBits(RA, 0, 1)
                .SetBits(0, 1, 3)
                .SetBits(Rcode, 4, 4);
            list.Add(b);

            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((Int16)Querys.Count)));
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((Int16)ResouceRecords.Count)));
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder(授权资源记录数)));
            list.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder(额外资源记录数)));

            foreach (var q in Querys)
            {
                list.AddRange(q.ToBytes());
            }
            foreach (var r in ResouceRecords)
            {
                list.AddRange(r.ToBytes());
            }

            return list.ToArray();
        }

        private int index;
        private readonly byte[] package;
        private byte ReadByte()
        {
            return package[index++];
        }
        private byte[] ReadBytes(int count = 1)
        {
            var bytes = new byte[count];
            for (var i = 0; i < count; i++)
                bytes[i] = ReadByte();
            return bytes;
        }

        public DNS(byte[] buffer, int length)
        {
            package = new byte[length];
            for (var i = 0; i < length; i++)
                package[i] = buffer[i];

            标志 = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReadBytes(2), 0));


            var b1 = ReadByte();
            var b2 = ReadByte();

            QR = b1.GetBits(0, 1);
            Opcode = b1.GetBits(1, 4);
            AA = b1.GetBits(5, 1);
            TC = b1.GetBits(6, 1);
            RD = b1.GetBits(7, 1);

            RA = b2.GetBits(0, 1);
            Rcode = b2.GetBits(4, 4);

            var queryCount = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReadBytes(2), 0));
            var rrCount = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReadBytes(2), 0));

            授权资源记录数 = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReadBytes(2), 0));
            额外资源记录数 = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(ReadBytes(2), 0));

            Querys = new List<Query>();
            for (var i = 0; i < queryCount; i++)
            {
                Querys.Add(new Query(ReadBytes));
            }

            for (var i = 0; i < rrCount; i++)
            {
                ResouceRecords.Add(new ResouceRecord(ReadBytes));
            }

        }
    }

    public static class Extension
    {
        public static int GetBits(this byte b, int start, int length)
        {
            var temp = b >> (8 - start - length);
            var mask = 0;
            for (var i = 0; i < length; i++)
            {
                mask = (mask << 1) + 1;
            }

            return temp & mask;

        }
        public static byte SetBits(this byte b, int data, int start, int length)
        {
            var temp = b;

            var mask = 0xFF;
            for (var i = 0; i < length; i++)
            {
                mask -= (0x01 << (7 - (start + i)));
            }
            temp = (byte)(temp & mask);

            mask = ((byte)data).GetBits(8 - length, length);
            mask <<= (7 - start);

            return (byte)(temp | mask);
        }
    }
}