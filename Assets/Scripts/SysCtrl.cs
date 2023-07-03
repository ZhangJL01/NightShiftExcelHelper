using NightShiftExcelHelper;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;
using UnityEngine.UI;

public class SysCtrl : MonoBehaviour
{
    [SerializeField] private Image codeImg;
    [SerializeField] private Button loadBtn;
    [SerializeField] private Text msgTxt;
    [SerializeField] private InputField codeIF;
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

    }

    private void FixedUpdate() {

        //msgTxt.textTxt.text = msgTxt.text;
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

        //NetHelper.Test();
        msgTxt.text = "信息安全相关网站测试中";

        yield return new WaitForFixedUpdate();

        msgTxt.text = "生成ERP弱口令表格";
        yield return new WaitForFixedUpdate();
        ExcelHelper.FillWPSheet(false);
        msgTxt.text = "生成RZOA弱口令表格";
        yield return null;
        ExcelHelper.FillWPSheet(true);
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
