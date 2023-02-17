using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class Test : MonoBehaviour
{
    private void Start()
    {
        Debug.Log(ExcelReaderManager.Instance.GetTable<Sheet1>().dataDic[1].name);
    }
    // Update is called once per frame
    void Update()
    {
        
    }
}
