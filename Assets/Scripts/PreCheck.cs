using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.SceneManagement;
using UnityEngine.UI;

namespace NightShiftExcelHelper {
    
    class PreCheck : MonoBehaviour {

        [SerializeField] private Button retry;
        [SerializeField] private Text note;

        private void Awake() {
            retry.onClick.AddListener(() => {
                Check();
            });

            Check();
        }
        private void Check() {
            if (DateTime.Now.Day.Equals(1)) {
                SceneManager.LoadScene(1);
            }

            if (File.GetLastWriteTime(ExcelHelper.nsSheet).DayOfYear < DateTime.Now.AddDays(-1).DayOfYear
                || File.GetLastWriteTime(ExcelHelper.wpSheet).DayOfYear < DateTime.Now.AddDays(-1).DayOfYear) {
                note.text = "更换模板文件为昨天！！！";
            }
            else {
                SceneManager.LoadScene(1);
            }
        }
    }
}

