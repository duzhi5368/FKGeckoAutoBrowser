using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using FKWebBrowser.Source;

namespace FKWebBrowser
{
    class ConnectHelper
    {
        private static bool CheckValideationResult(object sender, X509Certificate cert,
            X509Chain chain, SslPolicyErrors errors)
        {
            // 全部接受
            return true;
        }

        public static string UrlEncode(string str)
        {
            StringBuilder sb = new StringBuilder();
            byte[] byStr = Encoding.UTF8.GetBytes(str);
            for (int i = 0; i < byStr.Length; i++)
            {
                sb.Append(@"%" + Convert.ToString(byStr[i], 16));
            }
            return (sb.ToString());
        }
        // POST HTTP请求
        public static string HttpSendPost(string strPostUrl, string strPostDataStr, int nTimeOut, Log log = null)
        {
            string strEncodUrl = (strPostDataStr);
            if (strEncodUrl.Length >= 32)
            {
                if (log != null)
                    log.AddInfo(Log.ENUM_Level.Info, "Post url = " + strPostUrl + ", data = " + strEncodUrl.Substring(0, 32) + " ...");
                else
                    Console.WriteLine("Post url = " + strPostUrl + ", data = " + strEncodUrl.Substring(0, 32) + " ...");
            }
            else
            {
                if (log != null)
                    log.AddInfo(Log.ENUM_Level.Info, "Post url = " + strPostUrl);
                else
                    Console.WriteLine("Post url = " + strPostUrl);
            }
            ServicePointManager.DefaultConnectionLimit = 100;
            ServicePointManager.Expect100Continue = false;

            HttpWebRequest request = null;
            Stream requestStream = null;
            HttpWebResponse response = null;
            Stream responseStream = null;
            StreamReader streamReader = null;
            try
            {
                // https请求
                if (strPostUrl.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                {
                    ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValideationResult);
                    request = WebRequest.Create(strPostUrl) as HttpWebRequest;
                    request.ProtocolVersion = HttpVersion.Version10;
                }
                else
                {
                    request = (HttpWebRequest)(WebRequest.Create(new Uri(strPostUrl)));
                }
                byte[] byteArray = Encoding.UTF8.GetBytes(strEncodUrl);
                request.Timeout = nTimeOut;
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = byteArray.Length;
                requestStream = request.GetRequestStream();
                requestStream.Write(byteArray, 0, byteArray.Length);
                requestStream.Close();

                // 获取返回内容
                response = (HttpWebResponse)request.GetResponse();
                responseStream = response.GetResponseStream();
                streamReader = new StreamReader(responseStream, Encoding.GetEncoding("utf-8"));
                return streamReader.ReadToEnd();
            }
            catch (Exception e)
            {
                if (strEncodUrl.Length >= 32)
                {
                    if (log != null)
                        log.AddInfo(Log.ENUM_Level.Error, "[Error] Post url = " + strPostUrl + ", data = " + strEncodUrl.Substring(0, 32) + " ..." + ", ecxeption = " + e);
                    else
                        Console.WriteLine("[Error] Post url = " + strPostUrl + ", data = " + strEncodUrl.Substring(0, 32) + " ..." + ", ecxeption = " + e);
                }
                else
                {
                    if (log != null)
                        log.AddInfo(Log.ENUM_Level.Error, "[Error] Post url = " + strPostUrl  + ", ecxeption = " + e);
                    else
                        Console.WriteLine("[Error] Post url = " + strPostUrl + ", ecxeption = " + e);
                }
            }
            finally
            {
                if (streamReader != null)
                    streamReader.Close();
                if (responseStream != null)
                    responseStream.Close();
                if (response != null)
                    response.Close();
                if (request != null)
                    request.Abort();
            }
            return "";
        }

        // GET HTTP请求
        public static string HttpSendGet(string strGetUrl, int nTimeOut, Log log = null)
        {
            ServicePointManager.DefaultConnectionLimit = 100;
            HttpWebRequest request = null;
            HttpWebResponse response = null;
            Stream responseStream = null;
            StreamReader streamReader = null;
            try
            {
                if (strGetUrl.Length >= 32)
                {
                    if (log != null)
                        log.AddInfo(Log.ENUM_Level.Info, "Get url " + strGetUrl.Substring(0, 32) + " ...");
                    else
                        Console.WriteLine("Get url " + strGetUrl.Substring(0, 32) + " ...");
                }
                else
                {
                    if (log != null)
                        log.AddInfo(Log.ENUM_Level.Info, "Get url = " + strGetUrl);
                    else
                        Console.WriteLine("Get url = " + strGetUrl);
                }

                request = (HttpWebRequest)(WebRequest.Create(strGetUrl));
                request.Timeout = nTimeOut;
                request.Method = "GET";
                request.ContentType = "text/html;charset=UTF-8";

                // 获取返回内容
                response = (HttpWebResponse)request.GetResponse();
                responseStream = response.GetResponseStream();
                streamReader = new StreamReader(responseStream, Encoding.GetEncoding("utf-8"));
                return streamReader.ReadToEnd();
            }
            catch (Exception e)
            {
                if (strGetUrl.Length >= 32)
                {
                    if (log != null)
                        log.AddInfo(Log.ENUM_Level.Error, "[Error] Get url " + strGetUrl.Substring(0, 32) + "..." + ", ecxeption = " + e);
                    else
                        Console.WriteLine("[Error] Get url " + strGetUrl.Substring(0, 32) + "..." + ", ecxeption = " + e);
                }
                else
                {
                    if (log != null)
                        log.AddInfo(Log.ENUM_Level.Error, "[Error] Get url = " + strGetUrl);
                    else
                        Console.WriteLine("[Error] Get url = " + strGetUrl);
                }
            }
            finally
            {
                if (streamReader != null)
                    streamReader.Close();
                if (responseStream != null)
                    responseStream.Close();
                if (response != null)
                    response.Close();
                if (request != null)
                    request.Abort();
            }
            return "";
        }

        public static string GetLocalIPAddress()
        {
            var host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                    return ip.ToString();
            }
            throw new Exception("Local IP Address not found!");
        }


        public static string ToMD5Encrypt(string strSrcCode)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] result = md5.ComputeHash(Encoding.Default.GetBytes(strSrcCode));
            return BitConverter.ToString(result).Replace("-", "");

        }

        public static string ToJavaDesEncrypt(string strSrcCode, string strKey)
        {
            try
            {
                byte[] keyBytes = Encoding.UTF8.GetBytes(strKey);
                byte[] keyIV = keyBytes;
                byte[] inputByteArray = Encoding.UTF8.GetBytes(strSrcCode);

                // java 默认的是ECB模式，PKCS5padding；c#默认的CBC模式，PKCS7padding 所以这里我们默认使用ECB方式
                DESCryptoServiceProvider desProvider = new DESCryptoServiceProvider();
                desProvider.Mode = CipherMode.ECB;
                MemoryStream memStream = new MemoryStream();
                CryptoStream crypStream = new CryptoStream(memStream, desProvider.CreateEncryptor(keyBytes, keyIV), CryptoStreamMode.Write);
                crypStream.Write(inputByteArray, 0, inputByteArray.Length);
                crypStream.FlushFinalBlock();
                return Convert.ToBase64String(memStream.ToArray());
            }
            catch
            {
                return "";
            }
        }

        // 解密JAVA DES加密后的数据
        public static string FromJavaEncrypt(string decryptString, string key)
        {
            try
            {
                byte[] keyBytes = Encoding.UTF8.GetBytes(key);
                byte[] keyIV = keyBytes;
                byte[] inputByteArray = Convert.FromBase64String(decryptString);

                DESCryptoServiceProvider desProvider = new DESCryptoServiceProvider();

                // java 默认的是ECB模式，PKCS5padding；c#默认的CBC模式，PKCS7padding 所以这里我们默认使用ECB方式
                desProvider.Mode = CipherMode.ECB;
                MemoryStream memStream = new MemoryStream();
                CryptoStream crypStream = new CryptoStream(memStream, desProvider.CreateDecryptor(keyBytes, keyIV), CryptoStreamMode.Write);

                crypStream.Write(inputByteArray, 0, inputByteArray.Length);
                crypStream.FlushFinalBlock();
                return Encoding.Default.GetString(memStream.ToArray());
            }
            catch (Exception)
            {
                return "";
            }
        }
    }
}
