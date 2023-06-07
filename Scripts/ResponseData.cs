using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NightShiftExcelHelper {
    class ResponseData {

        //验证码
        public class Code {
            public string data;
        }

        //弱口令数据结构
        public class WPData {
            public Data data;

            public class Data {
                public List<WPInfo> data;
            }
            
            public class WPInfo {
                public string alarmName;                //事件名
                public string alarmResults;             //结果
                public string destAddress;              //目的IP
                public string srcAddress;               //来源IP
                public string srcUserName;              //用户名
                public string passwd;                   //密码
                public string collectorReceiptTime;     //时间
                public string destPort;                 //端口
                public string threatSeverity;           //威胁等级
                public string testResult;               //测试结果
                public string subCategory;              //告警子类型
            }
        }

        //风险资产数据结构
        public class RAData {
            public Data data;

            public class Data {
                public List<RAInfo> list;
            }

            public class RAInfo {
                public string ip;
                public string tags;
            }
        }

        //安全事件数据结构
        public class SIData {
            public List<SIInfo> data;

            public class SIInfo {
                public string incidentName;         //事件名称
                public Evidence evidence;           //信息
            }

            public class Evidence {
                public List<Body> body;
            }

            public class Body {
                public string srcAddress;
            }
        }
    }
}
