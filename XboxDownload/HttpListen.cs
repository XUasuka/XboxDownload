using System;
using System.Collections.Concurrent;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace XboxDownload
{
    class HttpListen
    {
        private readonly Form1 parentForm;
        private readonly ConcurrentDictionary<String, String> dicAppLocalUploadFile = new ConcurrentDictionary<String, String>();
        Socket socket = null;

        public HttpListen(Form1 parentForm)
        {
            this.parentForm = parentForm;
        }

        public void Listen()
        {
            IPEndPoint ipe = new IPEndPoint(Properties.Settings.Default.ListenIP == 0 ? IPAddress.Parse(Properties.Settings.Default.LocalIP) : IPAddress.Any, 80);
            socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            try
            {
                socket.Bind(ipe);
                socket.Listen(100);
            }
            catch (SocketException ex)
            {
                parentForm.Invoke(new Action(() =>
                {
                    parentForm.pictureBox1.Image = Properties.Resources.Xbox3;
                    MessageBox.Show(String.Format("启用HTTP服务失败!\n错误信息: {0}", ex.Message), "启用HTTP服务失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
                return;
            }
            while (Form1.bServiceFlag)
            {
                try
                {
                    Socket mySocket = socket.Accept();
                    ThreadPool.QueueUserWorkItem(TcpThread, mySocket);
                }
                catch { }
            }
        }

        private void TcpThread(object obj)
        {
            Socket mySocket = (Socket)obj;
            if (mySocket.Connected)
            {
                Byte[] _receive = new Byte[4096];
                int _num = mySocket.Receive(_receive, _receive.Length, 0);
                string _buffer = Encoding.ASCII.GetString(_receive, 0, _num);
                Match result = Regex.Match(_buffer, @"GET ([^\s]+)", RegexOptions.IgnoreCase);
                if (!result.Success)
                {
                    mySocket.Close();
                    return;
                }
                string _filePath = Regex.Replace(result.Groups[1].Value.Trim(), @"^http://[^/]+", "");
                result = Regex.Match(_buffer, @"Host:(.+)", RegexOptions.IgnoreCase);
                if (!result.Success)
                {
                    mySocket.Close();
                    return;
                }

                string _domainName = result.Groups[1].Value.Trim();
                string _localPath = null;
                if (Properties.Settings.Default.LocalUpload)
                {
                    if (File.Exists(Properties.Settings.Default.LocalPath + _filePath))
                        _localPath = Properties.Settings.Default.LocalPath + _filePath;
                    else if (File.Exists(Properties.Settings.Default.LocalPath + "\\" + Path.GetFileName(_filePath)))
                        _localPath = Properties.Settings.Default.LocalPath + "\\" + Path.GetFileName(_filePath);
                    else if (dicAppLocalUploadFile.ContainsKey(_filePath) && File.Exists(Properties.Settings.Default.LocalPath + "\\" + dicAppLocalUploadFile[_filePath]))
                        _localPath = Properties.Settings.Default.LocalPath + "\\" + dicAppLocalUploadFile[_filePath];
                }
                string _extension = Path.GetExtension(_filePath).ToLowerInvariant();
                if (Properties.Settings.Default.LocalUpload && !string.IsNullOrEmpty(_localPath))
                {
                    if (Form1.bRecordLog) parentForm.SaveLog("本地上传", _localPath, ((IPEndPoint)mySocket.RemoteEndPoint).Address.ToString());
                    using (FileStream fs = new FileStream(_localPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        using (BinaryReader br = new BinaryReader(fs))
                        {
                            string _contentRange = null, _status = "200 OK";
                            long _fileLength = br.BaseStream.Length, _startPosition = 0;
                            long _endPosition = _fileLength;
                            result = Regex.Match(_buffer, @"Range: bytes=(?<StartPosition>\d+)(-(?<EndPosition>\d+))?", RegexOptions.IgnoreCase);
                            if (result.Success)
                            {
                                _startPosition = long.Parse(result.Groups["StartPosition"].Value);
                                if (_startPosition > br.BaseStream.Length) _startPosition = 0;
                                if (!string.IsNullOrEmpty(result.Groups["EndPosition"].Value))
                                    _endPosition = long.Parse(result.Groups["EndPosition"].Value) + 1;
                                _contentRange = "bytes " + _startPosition + "-" + (_endPosition - 1) + "/" + _fileLength;
                                _status = "206 Partial Content";
                            }

                            StringBuilder sb = new StringBuilder();
                            sb.Append("HTTP/1.1 " + _status + "\r\n");
                            sb.Append("Server: Xbox-Skyer\r\n");
                            sb.Append("Content-Type: " + System.Web.MimeMapping.GetMimeMapping(_filePath) + "\r\n");
                            sb.Append("Content-Length: " + (_endPosition - _startPosition) + "\r\n");
                            if (_contentRange != null) sb.Append("Content-Range: " + _contentRange + "\r\n");
                            sb.Append("Accept-Ranges: bytes\r\n\r\n");

                            Byte[] _headers = Encoding.ASCII.GetBytes(sb.ToString());
                            mySocket.Send(_headers, 0, _headers.Length, SocketFlags.None, out _);

                            br.BaseStream.Position = _startPosition;
                            int _size = 4096;
                            while (Form1.bServiceFlag && mySocket.Connected)
                            {
                                long _remaining = _endPosition - br.BaseStream.Position;
                                if (Properties.Settings.Default.Truncation && _extension == ".xcp" && _remaining <= 1048576) //Xbox360主机本地上传防爆头
                                {
                                    Thread.Sleep(1000);
                                    continue;
                                }
                                byte[] _response = new byte[_remaining <= _size ? _remaining : _size];
                                br.Read(_response, 0, _response.Length);
                                mySocket.Send(_response, 0, _response.Length, SocketFlags.None, out _);
                                if (_remaining <= _size) break;
                            }
                        }
                    }
                }
                else
                {
                    bool _redirect = false;
                    string _cn = null;
                    if (Properties.Settings.Default.Redirect)
                    {
                        switch (_domainName)
                        {
                            case "assets1.xboxlive.com":
                            case "assets2.xboxlive.com":
                            case "dlassets.xboxlive.com":
                            case "dlassets2.xboxlive.com":
                            case "d1.xboxlive.com":
                            case "d2.xboxlive.com":
                                _redirect = true;
                                _cn = Regex.Replace(_domainName, @"\.com$", ".cn");
                                break;
                            case "7.assets1.xboxlive.com":
                            case "xvcf1.xboxlive.com":
                                _redirect = true;
                                _cn = "assets1.xboxlive.cn";
                                break;
                            case "xvcf2.xboxlive.com":
                                _redirect = true;
                                _cn = "assets2.xboxlive.cn";
                                break;
                        }
                    }
                    if (_redirect)
                    {
                        string _url = "http://" + _cn + _filePath;
                        if (Form1.bRecordLog) parentForm.SaveLog("HTTP 301", _url, ((IPEndPoint)mySocket.RemoteEndPoint).Address.ToString());
                        StringBuilder sb = new StringBuilder();
                        sb.Append("HTTP/1.1 301 Moved Permanently\r\n");
                        sb.Append("Server: Xbox-Skyer\r\n");
                        sb.Append("Content-Type: text/html\r\n");
                        sb.Append("Location: " + _url + "\r\n");
                        sb.Append("Content-Length: 0\r\n\r\n");
                        Byte[] _headers = Encoding.ASCII.GetBytes(sb.ToString());
                        mySocket.Send(_headers, 0, _headers.Length, SocketFlags.None, out _);
                    }
                    else
                    {
                        bool bFileNotFound = true;
                        string _url = "http://" + _domainName + _filePath;
                        if (_domainName == "dl.delivery.mp.microsoft.com" || ((_domainName == "download.xbox.com" || _domainName == "download.xbox.com.edgesuite.net" || _domainName == "xbox-ecn102.vo.msecnd.net") && (_extension == ".jpg" || _extension == ".png"))) //1.代理Xbox应用下载索引 2.代理Xbox360卖场图片
                        {
                            string ip = ClassWeb.HostToIP(_domainName, "114.114.114.114");
                            if (!string.IsNullOrEmpty(ip))
                            {
                                SocketPackage socketPackage = ClassWeb.HttpRequest(_url, "GET", null, null, true, false, false, null, null, null, null, null, null, null, null, 0, null, 30000, 30000, 1, ip);
                                if (string.IsNullOrEmpty(socketPackage.Err))
                                {
                                    if (Form1.bRecordLog) parentForm.SaveLog("HTTP 200", _url, ((IPEndPoint)mySocket.RemoteEndPoint).Address.ToString());
                                    bFileNotFound = false;
                                    StringBuilder sb = new StringBuilder();
                                    sb.Append("HTTP/1.1 200 OK\r\n");
                                    sb.Append("Server: Xbox-Skyer\r\n");
                                    sb.Append("Content-Type: " + System.Web.MimeMapping.GetMimeMapping(_filePath) + "\r\n");
                                    sb.Append("Content-Length: " + socketPackage.Buffer.Length + "\r\n\r\n");
                                    Byte[] _headers = Encoding.ASCII.GetBytes(sb.ToString());
                                    mySocket.Send(_headers, 0, _headers.Length, SocketFlags.None, out _);
                                    mySocket.Send(socketPackage.Buffer, 0, socketPackage.Buffer.Length, SocketFlags.None, out _);
                                }
                            }
                        }
                        else if (Properties.Settings.Default.LocalUpload && _domainName == "tlu.dl.delivery.mp.microsoft.com" && !dicAppLocalUploadFile.ContainsKey(_filePath)) //识别本地上传应用文件名
                        {
                            string ip = ClassWeb.HostToIP(_domainName, "114.114.114.114");
                            if (!string.IsNullOrEmpty(ip))
                            {
                                SocketPackage socketPackage = ClassWeb.HttpRequest(_url, "GET", null, null, true, false, false, null, null, new string[] { "Range: bytes=0-0" }, null, null, null, null, null, 0, null, 30000, 30000, 1, ip);
                                Match m1 = Regex.Match(socketPackage.Headers, @"Content-Disposition: attachment; filename=(.+)");
                                if (m1.Success)
                                {
                                    string filename = m1.Groups[1].Value.Trim();
                                    dicAppLocalUploadFile.AddOrUpdate(_filePath, filename, (oldkey, oldvalue) => filename);
                                }
                            }
                        }
                        if (bFileNotFound)
                        {
                            if (Form1.bRecordLog) parentForm.SaveLog("HTTP 404", _url, ((IPEndPoint)mySocket.RemoteEndPoint).Address.ToString());
                            Byte[] _response = Encoding.ASCII.GetBytes("File not found.");
                            StringBuilder sb = new StringBuilder();
                            sb.Append("HTTP/1.1 404 Not Found\r\n");
                            sb.Append("Server: Xbox-Skyer\r\n");
                            sb.Append("Content-Type: text/html\r\n");
                            sb.Append("Content-Length: " + _response.Length + "\r\n\r\n");
                            Byte[] _headers = Encoding.ASCII.GetBytes(sb.ToString());
                            mySocket.Send(_headers, 0, _headers.Length, SocketFlags.None, out _);
                            mySocket.Send(_response, 0, _response.Length, SocketFlags.None, out _);
                        }
                    }
                }
                try
                {
                    mySocket.Shutdown(SocketShutdown.Both);
                }
                catch { }
                mySocket.Close();
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
}
