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
            //EditorUtility.DisplayDialog("?...", e.Message, "ȷ��");
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
            .Replace("data:image/jpeg;base64,", "");//��base64ͷ����Ϣ�滻
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


            //����Web���ʶ���
            myRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create("http://rzoa.sdgt.com/names.nsf?Login");
            cc = new System.Net.CookieContainer();

            id = UnityWebRequest.EscapeURL("����ΰ");
            password = UnityWebRequest.EscapeURL("710119");

            cc.Add(new Cookie("myusername", id, "/", "rzoa.sdgt.com"));
            cc.Add(new Cookie("indi_locale", "zh-cn", "/", ".sdgt.com"));
            myRequest.Method = "POST";
            myRequest.ContentType = "application/x-www-form-urlencoded";
            myRequest.AllowAutoRedirect = true;

            myRequest.CookieContainer = cc;
            //���û�����ת�ɡ�UTF-8�����ֽ���
            #region ���Post ����

            s = "Username=" + id + "&Password=" + password;
        Debug.Log(s);
            myRequest.ContentLength = s.Length;
            using (reqStream = new StreamWriter(myRequest.GetRequestStream())) {
                reqStream.Write(s);
                reqStream.Close();
            }
            #endregion

            //��ȡ�ӿڷ���ֵ
            //ͨ��Web���ʶ����ȡ��Ӧ����
            myResponse = (HttpWebResponse)myRequest.GetResponse();
            //ͨ����Ӧ����������StreamReader������ΪStreamReader���߼�����
            reader = new StreamReader(myResponse.GetResponseStream(), System.Text.Encoding.GetEncoding("GB2312"));
            //string returnXml = UnityWebRequest.EscapeURL(reader.ReadToEnd());//����б�����������������
            returnXml = reader.ReadToEnd();//����StreamReader�Ϳ��Դ���Ӧ���ݴ�ͷ����β
            reader.Close();
            myResponse.Close();
         msgTxt.text = returnXml.Contains("<title>�칫�Զ���ϵͳ</title>") ? "�ɹ�" : "ʧ��";


    }

    private IEnumerator Check() {

        if (codeIF.text.Trim() == "") {
            yield break;
        }
        msgTxt.text = "��¼��";
        yield return new WaitForEndOfFrame();


        if (NetHelper.Login(codeIF.text.Trim())) {
            codeImg.gameObject.SetActive(false);
            codeIF.gameObject.SetActive(false);
            loadBtn.gameObject.SetActive(false);
            msgTxt.gameObject.SetActive(true);
        } else {
            msgTxt.text = "��¼ʧ��";
            yield break;
        }
        msgTxt.text = "��֤��";
        yield return null;
        NetHelper.Authentication();

        msgTxt.text = "��������Ϣͳ���У�RZOA";
        yield return new WaitForEndOfFrame();
        NetHelper.GetWPList("10.200.24.11");

        msgTxt.text = "��������Ϣͳ���У�ERP";
        yield return null;
        NetHelper.GetWPList("10.200.35.56");


        msgTxt.text = "��ȫ�¼�ͳ���У���Σ";
        yield return null;
        NetHelper.GetSIList("High");

        msgTxt.text = "��ȫ�¼�ͳ���У���Σ";
        yield return null;
        NetHelper.GetSIList("Medium");

        msgTxt.text = "��ȫ�¼�ͳ���У���Σ";
        yield return null;
        NetHelper.GetSIList("Low");

        msgTxt.text = "�����ʲ�ͳ���У���ʧ��";
        yield return null;
        NetHelper.GetRAList();

        msgTxt.text = "��������Ϣ�����У�RZOA";
        yield return null;
        NetHelper.RzoaRequset();

        msgTxt.text = "��������Ϣ�����У�ERP";
        yield return null;
        NetHelper.ErpRequset();

        NetHelper.Test();
        msgTxt.text = "��Ϣ��ȫ�����վ������";

        yield return new WaitForFixedUpdate();

        msgTxt.text = "������������";
        yield return null;
        ExcelHelper.FillWPSheet();
        msgTxt.text = "������Ϣ��ȫ���";
        yield return null;
        ExcelHelper.FillNSSheet();

        msgTxt.text = "��ɣ�\n����͵���ɹ���һ����";
        yield return null;
        Invoke("Exit", 10000);
    }

    void Exit() {
        Application.Quit();
    }
}
