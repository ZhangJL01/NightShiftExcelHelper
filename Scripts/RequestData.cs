using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NightShiftExcelHelper {
    class RequestData {

        //登录数据结构
        public class LoginData {
            public string type = "password";
            public TypeData typeData = new TypeData();
            public class TypeData {
                public string username = "admin";
                public string password = "f657d529d415914514f4d40a5f788ade97b1f134e4e093baedc55da81ca3b119";
                public string otherPassword = "ceb7b0f380c0eb555cda5d2138c26406b40f5bd9";
                public string code;
            }
        }

        //弱口令数据结构
        public class WPData {

            public string queryStr = "destAddress ==";
            public Condition condition = new Condition();

            public string startTime = String.Format("{0:yyyy'-'MM'-'dd' 'HH':'00':'00}", DateTime.Now.AddDays(-1));
            public string endTime = String.Format("{0:yyyy'-'MM'-'dd' 'HH':'mm':'ss}", DateTime.Now);
            public int from = 0;
            public int size = 5000;
            public int searchTypeNum = 3;
            public string url = "alarms";
            public int pageSize = 500;
            public string scene = "自由探索";


            public class Condition {
                public Query query = new Query();
            }

            public class Query {
                [JsonProperty("bool")]
                public Bool b = new Bool();
            }

            public class Bool {
                public Filter filter = new Filter();
            }

            public class Filter {
                public Term term = new Term();
            }

            public class Term {
                public string destAddress = "10.200.24.11";
            }

        }

        //风险资产数据结构
        //public class RAData {
        //    //offset=0&limit=10&riskyType=fallen&orgId=&keyword=&date=2023-06-04
        //    public int offset = 0;
        //    public int limit = 10;
        //    //public string riskyType = "";
        //    public string orgId = "";
        //    public string keyword = "";
        //    public string date = String.Format("{0:yyyy'-'MM'-'dd}", DateTime.Now.AddDays(-1));
        //}

    }
}
