using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.VisualBasic.FileIO;

namespace NCreoConvertTool
{
    class MyUtils
    {
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        [DllImport("kernel32")]
        private static extern int WritePrivateProfileString(string section, string key, string val, string filePath);

        public static string ReadIniKeyValue(string iniPath, string section, string key, string defaultValue)
        {
            const int MAX_BUFFER = 65535;
            var sb = new StringBuilder(MAX_BUFFER);
            GetPrivateProfileString(section, key, defaultValue, sb, MAX_BUFFER, iniPath);
            return sb.ToString();
        }

        public static void WriteIniKeyValue(string iniPath, string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, iniPath);
        }

        public static string GBK2UTF8(string gbkString)
        {
            byte[] gbkBytes = Encoding.Default.GetBytes(gbkString);

            // 如果系统默认不是GBK，可以使用GBK编码获取字节
            // byte[] gbkBytes = Encoding.GetEncoding("GBK").GetBytes(gbkString);

            byte[] utf8Bytes = Encoding.Convert(Encoding.GetEncoding("GBK"), Encoding.UTF8, gbkBytes);
            string utf8String = Encoding.UTF8.GetString(utf8Bytes);
            return utf8String;
        }

        public static string UTF82GBK(string utf8String)
        {
            byte[] utf8Bytes = Encoding.UTF8.GetBytes(utf8String);
            string gbkString = Encoding.GetEncoding("GBK").GetString(utf8Bytes);
            return gbkString;
        }

        public static string CalculateMD5(string filePath)
        {
            using (var md5 = new MD5CryptoServiceProvider())
            {
                using (var stream = File.OpenRead(filePath))
                {
                    byte[] hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLower();
                }
            }
        }

        /// <summary>
        /// 生成MD5码版本号:读取目前软件执行档然后产生MD5码，作为软件版本。
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        public static string GetFileMd5Code(string filePath)
        {
            System.Text.StringBuilder builder = new System.Text.StringBuilder();
            if (filePath != "")
            {
                using (var md5 = new System.Security.Cryptography.MD5CryptoServiceProvider())
                {
                    string backfilename = filePath + "e";
                    if (System.IO.File.Exists(backfilename) == true)
                    {
                        System.IO.File.Delete(backfilename);
                    }

                    System.IO.File.Copy(filePath, backfilename);//复制一份，防止占用

                    //利用复制的执行档建立MD5码
                    using (System.IO.FileStream fs = new System.IO.FileStream(filePath + "e", FileMode.Open))
                    {
                        byte[] bt = md5.ComputeHash(fs);
                        for (int i = 0; i < bt.Length; i++)
                        {
                            builder.Append(bt[i].ToString("x2"));
                        }
                    }
                    System.IO.File.Delete(filePath + "e");//删除复制的文件，这里没处理异常.
                }
            }

            return builder.ToString();
        }
    }
}
