using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class TestExample : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        Debug.Log(ExcelReaderManager.Instance.GetTable<Sheet1>().dataDic[1].atk);//��ȡ��sheet1�� ��1�����ݵ�atk ע����sheet1�Ǳ��� ����Excel�ļ�����

    }

    // Update is called once per frame
    void Update()
    {
        
    }
}
