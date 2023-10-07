using NightShiftExcelHelper;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using UnityEditor;
using UnityEngine;
using UnityEngine.Networking;
using UnityEngine.UI;

public class SysCtrl : MonoBehaviour
{
    [SerializeField] private Image codeImg;
    [SerializeField] private Button loadBtn;
    [SerializeField] private Text msgTxt;
    [SerializeField] private InputField codeIF;
    //[SerializeField] private Button test;
    //[SerializeField] public Text path;


    private void Awake() {
        try {
            StartCoroutine(GetCode());
            loadBtn.onClick.AddListener(() => {
                StartCoroutine(Check());
            });

        } catch (Exception e) {
            //EditorUtility.DisplayDialog("?...", e.Message, "确定");
        }

        //test.onClick.AddListener(() => {
        //     RzoaRequset();
        //});
    }


    private IEnumerator GetCode() {
        string base64 = NetHelper.GetCode();
        yield return null;
        base64 = base64.Replace("data:image/png;base64,", "")
            .Replace("data:image/jgp;base64,", "").Replace("data:image/jpg;base64,", "")
            .Replace("data:image/jpeg;base64,", "");//将base64头部信息替换
        byte[] bytes = Convert.FromBase64String(base64);

        Texture2D texture = new Texture2D(200, 100);
        texture.LoadImage(bytes);
        Sprite sprite = Sprite.Create(texture, new Rect(0, 0, texture.width, texture.height), new Vector2(0.5f, 0.5f));

        codeImg.sprite = sprite;

        //Username % 3d % e5 % ae % 8b % e5 % 85 % 86 % e4 % bc % 9f % 26Password % 3d710119
        //"%e5%ae%8b%e5%85%86%e4%bc%9f"
    }

    public void RzoaRequset() {

        string id, password, s, returnXml;

        System.Net.HttpWebRequest myRequest;
        System.Net.CookieContainer cc;
        System.Net.HttpWebResponse myResponse;
        StreamReader reader;
        StreamWriter reqStream;


            //创建Web访问对象
            myRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create("http://rzoa.sdgt.com/names.nsf?Login");
            cc = new System.Net.CookieContainer();

            id = UnityWebRequest.EscapeURL("宋兆伟");
            password = UnityWebRequest.EscapeURL("710119");

            cc.Add(new Cookie("myusername", id, "/", "rzoa.sdgt.com"));
            cc.Add(new Cookie("indi_locale", "zh-cn", "/", ".sdgt.com"));
            myRequest.Method = "POST";
            myRequest.ContentType = "application/x-www-form-urlencoded";
            myRequest.AllowAutoRedirect = true;

            myRequest.CookieContainer = cc;
            //把用户数据转成“UTF-8”的字节流
            #region 添加Post 参数

            s = "Username=" + id + "&Password=" + password;
        Debug.Log(s);
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
            reader = new StreamReader(myResponse.GetResponseStream(), System.Text.Encoding.GetEncoding("GB2312"));
            //string returnXml = UnityWebRequest.EscapeURL(reader.ReadToEnd());//如果有编码问题就用这个方法
            returnXml = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
            reader.Close();
            myResponse.Close();
         msgTxt.text = returnXml.Contains("<title>办公自动化系统</title>") ? "成功" : "失败";


    }

    private IEnumerator Check() {

        if (codeIF.text.Trim() == "") {
            yield break;
        }
        msgTxt.text = "登录中";
        yield return new WaitForEndOfFrame();


        if (NetHelper.Login(codeIF.text.Trim())) {
            codeImg.gameObject.SetActive(false);
            codeIF.gameObject.SetActive(false);
            loadBtn.gameObject.SetActive(false);
            msgTxt.gameObject.SetActive(true);
        } else {
            msgTxt.text = "登录失败";
            yield break;
        }
        msgTxt.text = "认证中";
        yield return null;
        NetHelper.Authentication();

        msgTxt.text = "弱口令信息统计中：RZOA";
        yield return new WaitForEndOfFrame();
        NetHelper.GetWPList("10.200.24.11");

        msgTxt.text = "弱口令信息统计中：ERP";
        yield return null;
        NetHelper.GetWPList("10.200.35.56");


        msgTxt.text = "安全事件统计中：高危";
        yield return null;
        NetHelper.GetSIList("High");

        msgTxt.text = "安全事件统计中：中危";
        yield return null;
        NetHelper.GetSIList("Medium");

        msgTxt.text = "安全事件统计中：低危";
        yield return null;
        NetHelper.GetSIList("Low");

        msgTxt.text = "风险资产统计中：已失陷";
        yield return null;
        NetHelper.GetRAList();

        msgTxt.text = "弱口令信息测试中：RZOA";
        yield return null;
        NetHelper.RzoaRequset();

        msgTxt.text = "弱口令信息测试中：ERP";
        yield return null;
        NetHelper.ErpRequset();

        NetHelper.Test();
        msgTxt.text = "信息安全相关网站测试中";

        yield return new WaitForFixedUpdate();

        msgTxt.text = "生成弱口令表格";
        yield return null;
        ExcelHelper.FillWPSheet();
        msgTxt.text = "生成信息安全表格";
        yield return null;
        ExcelHelper.FillNSSheet();

        msgTxt.text = "完成！\n又是偷懒成功的一天呢";
        yield return null;
        Invoke("Exit", 10000);
    }

    void Exit() {
        Application.Quit();
    }
}
