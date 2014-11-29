using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Threading;
using System.Web;
using System.Net;
using System.Diagnostics;

namespace ConsoleApplication7
{
    class Program
    {






        public static string TextToBase64(string sAscii)
        {
            System.Text.ASCIIEncoding encoding = new System.Text.ASCIIEncoding();
            byte[] bytes = encoding.GetBytes(sAscii);
            return System.Convert.ToBase64String(bytes, 0, bytes.Length);
        }



        private static string CheckMail()
        {
            string result = "0";

            try
            {
                var url = @"https://mail.google.com/mail/feed/atom";
                var USER = "your email@gmail.com";
                var PASS = "your password";
            // gmail set to less secure apps access https://www.google.com/settings/security/lesssecureapps //
                var encoded = TextToBase64(USER + ":" + PASS);

                var myWebRequest = HttpWebRequest.Create(url);
                myWebRequest.Method = "POST";
                myWebRequest.ContentLength = 0;
                myWebRequest.Headers.Add("Authorization", "Basic " + encoded);

                var response = myWebRequest.GetResponse();
                var stream = response.GetResponseStream();

                XmlReader reader = XmlReader.Create(stream);
                while (reader.Read())
                    if (reader.NodeType == XmlNodeType.Element)
                        if (reader.Name == "fullcount")
                        {
                            result = reader.ReadElementContentAsString();
                            return result;
                        }
            }
            catch (Exception ee) { Console.WriteLine(ee.Message); }
            return result;

        }

        [STAThread]
        static void Main(string[] args)
        {
            string targetDir = string.Format(@"D:\");//this is where mybatch.bat lies


            Process scriptProc = new Process();
            scriptProc.StartInfo.WorkingDirectory = targetDir;

            scriptProc.StartInfo.FileName = "a.vbs";
            //scriptProc.StartInfo.Arguments = "//B //Nologo D:\a.vbs";
            scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;

            //scriptProc.Start();
            // scriptProc.WaitForExit();
            // scriptProc.Close();

            try
            {
                //SerialPort port = new SerialPort("COM9", 9600, Parity.None, 8, StopBits.One);
                // port.Open();

                string Unreadz = "0";
                while (true)
                {

                    Unreadz = CheckMail();
                    Console.WriteLine("Unread Mails: " + Unreadz);

                    if (!Unreadz.Equals("0")) { scriptProc.Start(); Thread.Sleep(10000); scriptProc.WaitForExit(); /*scriptProc.Close();*/ }
                    else Console.WriteLine("no emails");















                    Thread.Sleep(10000);
                }
            }
            catch (Exception ee) { Console.WriteLine(ee.Message); }



        }
    }
}


