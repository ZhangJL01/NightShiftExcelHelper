using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace NightShiftExcelHelper {
    class NetHelper {

        private static string mainUrl = "https://10.200.6.170/";
        private static string codeUrl = "https://10.200.6.170/api/v1.0/verify/code";
        private static string loginUrl = "https://10.200.6.170/api/v1.0/login";
        private static string authorUrl =
            "https://10.200.6.170/api/authentication/forPassword?otherPassword=ceb7b0f380c0eb555cda5d2138c26406b40f5bd9&password=f657d529d415914514f4d40a5f788ade97b1f134e4e093baedc55da81ca3b119&cascadeOrgId=fe22d914-4ad0-11ed-96f9-024248760126_1665649699&connectType=ALL";
        private static string wpUrl =
            "https://10.200.6.170/api/retrieve/alarms/getDataList?cascadeOrgId=fe22d914-4ad0-11ed-96f9-024248760126_1665649699&connectType=ALL";
        private static string raUrl =
            "https://10.200.6.170/api/alarmCenter/listRiskyAssets?cascadeOrgId=fe22d914-4ad0-11ed-96f9-024248760126_1665649699&connectType=ALL";
        private static string siUrl =
            "https://10.200.6.170/api/security/incident/fuzzyQuery?keyword=&startTime={0}&endTime={1}&offset=0&limit=50&threatSeverity={2}&statuses=unprocessed&cascadeOrgId=fe22d914-4ad0-11ed-96f9-024248760126_1665649699&connectType=ALL";

        private static string erpUrl = "http://erpapp.sdgtrz.com:8088/OA_HTML/AppsLocalLogin.jsp";
        private static string rzoaUrl = "http://rzoa.sdgt.com/names.nsf?Login";

        private static Cookie cookie = new Cookie();

        public static List<ResponseData.WPData.WPInfo> rzoaList;
        public static List<ResponseData.WPData.WPInfo> erpList;

        public static List<ResponseData.RAData.RAInfo> inNetList;
        public static List<ResponseData.RAData.RAInfo> exNetList;
        public static List<ResponseData.RAData.RAInfo> icNetList;
        public static List<ResponseData.RAData.RAInfo> otherNetList;

        public static List<ResponseData.SIData.SIInfo> highSIList;
        public static List<ResponseData.SIData.SIInfo> mediumSIList;
        public static List<ResponseData.SIData.SIInfo> lowSIList;

        private static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; }

        public static void GetCookie() {

            ServicePointManager.ServerCertificateValidationCallback = ValidateServerCertificate;

            HttpWebRequest myRequest = WebRequest.CreateHttp(mainUrl);

            myRequest.Method = "GET";
            myRequest.ContentType = "application/json";
            //myRequest.Headers.Add("X-Service", "AuthenticateUser");

            //获取接口返回值
            //通过Web访问对象获取响应内容
            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();

            //不知道为啥会被清空引用
            string strCookie = String.Copy(myResponse.Headers["Set-Cookie"]);
            StrToCookie(strCookie);
            myResponse.Close();
        }

        public static string GetCode() {
            GetCookie();
            HttpWebRequest myRequest = WebRequest.CreateHttp(codeUrl);
            myRequest.Method = "POST";
            myRequest.ContentType = "application/json";
            myRequest.Referer = mainUrl;

            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cookie);
            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();

            //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
            StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
            //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
            string returnStr = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
            reader.Close();

            //string strCookie = String.Copy(myResponse.Headers["Set-Cookie"]);
            //StrToCookie(strCookie);
            myResponse.Close();
            return JsonConvert.DeserializeObject<ResponseData.Code>(returnStr).data;
        }

        public static bool Login(string code) {
            HttpWebRequest myRequest = WebRequest.CreateHttp(loginUrl);
            myRequest.Method = "POST";
            myRequest.ContentType = "application/json";
            myRequest.Referer = mainUrl;
            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cookie);

            #region 添加Post 参数
            RequestData.LoginData loginData = new RequestData.LoginData();
            loginData.typeData.code = code;
            string s = JsonConvert.SerializeObject(loginData);
            byte[] data = Encoding.Default.GetBytes(s);
            myRequest.ContentLength = s.Length;
            using (var reqStream = new StreamWriter(myRequest.GetRequestStream())) {
                reqStream.Write(s);
                reqStream.Close();
            }
            #endregion      

            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
            //HttpContext.Current.Request.Cookies.Add(cookie);
            //myRequest.CookieContainer.
            //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
            StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
            //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
            string returnStr = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
            reader.Close();

            string strCookie = String.Copy(myResponse.Headers["Set-Cookie"]);
            StrToCookie(strCookie);
            myResponse.Close();

            return myResponse.StatusCode == HttpStatusCode.OK ? true : false;
        }

        public static void Authentication() {
            HttpWebRequest myRequest = WebRequest.CreateHttp(authorUrl);
            myRequest.Method = "GET";
            myRequest.ContentType = "application/json";
            myRequest.Referer = mainUrl;
            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cookie);


            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
            myResponse.Close();
        }

        public static void GetWPList(string destAdd) {
            HttpWebRequest myRequest = WebRequest.CreateHttp(wpUrl);
            myRequest.Method = "POST";
            myRequest.ContentType = "application/json";
            myRequest.Referer = mainUrl;
            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cookie);

            #region 添加Post 参数
            RequestData.WPData wpData = new RequestData.WPData();
            wpData.condition.query.b.filter.term.destAddress = destAdd;
            wpData.queryStr += " \"" + destAdd + "\" ";
            wpData.startTime = String.Format("{0:yyyy'-'MM'-'dd' 'HH':'00':'00}", DateTime.Now.AddDays(-1));
            wpData.endTime = String.Format("{0:yyyy'-'MM'-'dd' 'HH':'mm':'ss}", DateTime.Now);
            string s = JsonConvert.SerializeObject(wpData);
            byte[] data = Encoding.Default.GetBytes(s);
            //myRequest.ContentLength = data.Length;
            myRequest.SendChunked = true;
            using (var reqStream = new StreamWriter(myRequest.GetRequestStream())) {
                reqStream.Write(s);
                reqStream.Close();
            }
            #endregion      

            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
            //HttpContext.Current.Request.Cookies.Add(cookie);
            //myRequest.CookieContainer.
            //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
            StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
            //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
            string returnStr = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
            reader.Close();
            if (destAdd == "10.200.24.11") {
                rzoaList = JsonConvert.DeserializeObject<ResponseData.WPData>(returnStr).data.data;
            }
            else {
                erpList = JsonConvert.DeserializeObject<ResponseData.WPData>(returnStr).data.data;
            }
            myResponse.Close();
        }

        public static void GetRAList() {
            HttpWebRequest myRequest = WebRequest.CreateHttp(raUrl);
            myRequest.Method = "POST";
            myRequest.ContentType = "application/json";
            myRequest.Referer = mainUrl;
            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cookie);

            #region 添加Post 参数
            RequestData.RAData raData = new RequestData.RAData();
            string s = JsonConvert.SerializeObject(raData);
            //myRequest.ContentLength = data.Length;
            myRequest.SendChunked = true;
            using (var reqStream = new StreamWriter(myRequest.GetRequestStream())) {
                reqStream.Write(s);
                reqStream.Close();
            }
            #endregion      

            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
            //HttpContext.Current.Request.Cookies.Add(cookie);
            //myRequest.CookieContainer.
            //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
            StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
            //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
            string returnStr = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
            reader.Close();

            List<ResponseData.RAData.RAInfo> raList = 
                JsonConvert.DeserializeObject<ResponseData.RAData>(returnStr).data.list;

            foreach(ResponseData.RAData.RAInfo ra in raList) {
                if (ra.assetName.Contains("内网")) {
                    inNetList.Add(ra);
                } else if (ra.assetName.Contains("外网")) {
                    exNetList.Add(ra);
                } else if (ra.assetName.Contains("工控网")){
                    icNetList.Add(ra);
                } else {
                    otherNetList.Add(ra);
                }
                
            }  

            myResponse.Close();
        }

        public static void GetSIList(string threadLevel) {
            //MediumHighLow
            string startTime = 
                HttpUtility.UrlEncode(String.Format("{0:yyyy'-'MM'-'dd' '00':'00':'00}", DateTime.Now.AddDays(-1)));
            string endTime =
                HttpUtility.UrlEncode(String.Format("{0:yyyy'-'MM'-'dd' '23':'59':'59}", DateTime.Now.AddDays(-1)));

            HttpWebRequest myRequest = WebRequest.CreateHttp(String.Format(siUrl, startTime, endTime, threadLevel));
            myRequest.Method = "GET";
            myRequest.ContentType = "application/json";
            myRequest.Referer = mainUrl;
            myRequest.CookieContainer = new CookieContainer();
            myRequest.CookieContainer.Add(cookie);     

            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
            //HttpContext.Current.Request.Cookies.Add(cookie);
            //myRequest.CookieContainer.
            //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
            StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
            //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
            string returnStr = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
            reader.Close();

            switch(threadLevel) {
                case "High":
                    highSIList = JsonConvert.DeserializeObject<ResponseData.SIData>(returnStr).data;
                    break;
                case "Medium":
                    mediumSIList = JsonConvert.DeserializeObject<ResponseData.SIData>(returnStr).data;
                    break;
                case "Low":
                    lowSIList = JsonConvert.DeserializeObject<ResponseData.SIData>(returnStr).data;
                    break;
                default:
                    break;
            }

            myResponse.Close();
        }

        public static void ErpRequset() {

            string a, returnXml;
            HttpWebRequest myRequest;
            HttpWebResponse myResponse;
            StreamReader reader;

            foreach (ResponseData.WPData.WPInfo wpInfo in erpList) {
                a = erpUrl + "?username=" + wpInfo.srcUserName + "&password=" + wpInfo.passwd;

                //创建Web访问对象
                myRequest = (HttpWebRequest)WebRequest.Create(a);

                myRequest.Method = "POST";
                myRequest.ContentType = "application/json";
                myRequest.Headers.Add("X-Service", "AuthenticateUser");

                //获取接口返回值
                //通过Web访问对象获取响应内容
                myResponse = (HttpWebResponse)myRequest.GetResponse();
                //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
                reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
                //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
                returnXml = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
                reader.Close();
                myResponse.Close();
                wpInfo.testResult = returnXml.Contains("status: 'success'") ? "成功" : "失败";
            }
                       
        }

        public static void RzoaRequset() {

            string id, password, s, returnXml;
            byte[] data;

            HttpWebRequest myRequest;
            CookieContainer cc;
            HttpWebResponse myResponse;
            StreamReader reader;
            StreamWriter reqStream;

            foreach (ResponseData.WPData.WPInfo wpInfo in rzoaList) {
                //创建Web访问对象
                myRequest = (HttpWebRequest)WebRequest.Create(rzoaUrl);
                cc = new CookieContainer();

                id = HttpUtility.UrlEncode(wpInfo.srcUserName);
                password = HttpUtility.UrlEncode(wpInfo.passwd);

                cc.Add(new Cookie("myusername", id, "/", "rzoa.sdgt.com"));
                cc.Add(new Cookie("indi_locale", "zh-cn", "/", ".sdgt.com"));
                myRequest.Method = "POST";
                myRequest.ContentType = "application/x-www-form-urlencoded";
                myRequest.AllowAutoRedirect = true;

                myRequest.CookieContainer = cc;
                //把用户数据转成“UTF-8”的字节流
                #region 添加Post 参数
                
                s = "Username=" + id + "&Password=" + password;
                data = Encoding.Default.GetBytes(HttpUtility.UrlEncode(s));
                myRequest.ContentLength = s.Length;
                using (reqStream = new StreamWriter(myRequest.GetRequestStream())) {
                    reqStream.Write(s);
                    reqStream.Close();
                }
                #endregion

                //获取接口返回值
                //通过Web访问对象获取响应内容
                myResponse = (HttpWebResponse)myRequest.GetResponse();
                //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
                reader = new StreamReader(myResponse.GetResponseStream(), Encoding.Default);
                //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
                returnXml = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
                reader.Close();
                myResponse.Close();
                wpInfo.testResult = returnXml.Contains("<title>办公自动化系统</title>") ? "成功" : "失败";
            }
            
        }

        public static void StrToCookie(string strCookie) {
            //"ailpha-bigdata=NzIzN2ZmYWMtNzZlZS00YTZjLWExZDktYTQ0YTdiMmRiYjg4; Path=/; Secure; HttpOnly; SameSite=None"
            string te = strCookie.Split(';')[0];
            
            cookie.Name = te.Split('=')[0];
            cookie.Value = te.Split('=')[1];
            cookie.Path = "/";
            cookie.Secure = true;
            cookie.HttpOnly = true;
            cookie.Domain = "10.200.6.170";
        }

        public static void Test() {
            Console.WriteLine("wancheng");
        }
    }
}
