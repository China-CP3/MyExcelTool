using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class TestExample : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        Debug.Log(ExcelReaderManager.Instance.GetTable<Sheet1>().dataDic[1].atk);//读取表sheet1中 第1行数据的atk 注意是sheet1是表名 不是Excel文件名！

    }

    // Update is called once per frame
    void Update()
    {
        
    }
}
