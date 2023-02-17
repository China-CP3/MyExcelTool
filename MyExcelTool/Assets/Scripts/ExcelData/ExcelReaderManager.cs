using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using UnityEngine;

public class ExcelReaderManager
{
    private static ExcelReaderManager instance;
    public static ExcelReaderManager Instance
    {
        get
        {
            if (instance == null)
            {
                instance = new ExcelReaderManager();
            }
            return instance;
        }
    }

    public static string BinaryFile_Path = Application.streamingAssetsPath + "/BinaryFile/";
    public static string fieldClassPath = Application.dataPath + "/Scripts/ExcelData/fieldClass/";//���ɵ��ֶ��ļ����·��
    public static string dicClassPath = Application.dataPath + "/Scripts/ExcelData/dicClass/";//���ɵ��ֵ��ļ����·��
    public static string ExcelFilePath = Application.dataPath + "/Editor/Excel/ExcelFile/";//excel�ļ����·��

    private Dictionary<string, object> tableDic = new Dictionary<string, object>();

    public  ExcelReaderManager()
    {
        InitDate();
    }
    public void InitDate()
    {
        //if (!Directory.Exists(dicClassPath))
        //{
        //    Directory.CreateDirectory(dicClassPath);
        //}

        //DirectoryInfo directory = new DirectoryInfo(dicClassPath);
        //FileInfo[] files = directory.GetFiles();
        //for (int i = 0; i < files.Length; i++)
        //{
        //    if(files[i].Extension==".cs")
        //    {
        //        string className = files[i].Name.Remove(files[i].Name.Length - 3);
        //        Type  t = Type.GetType(className);

        //    }         
        //}
        LoadTable<Sheet1, Sheet1fieldClass>();
    }
    /// <summary>
    /// ����Excel���2�������ݵ��ֵ���
    /// </summary>
    /// <typeparam name="dicClass">�ֵ���</typeparam>
    /// <typeparam name="FieldClass">�ֶ���</typeparam>
    private void LoadTable<dicClass,FieldClass>()
    {
        //��ȡ excel���Ӧ��2�����ļ� �����н���
        using (FileStream fs=new FileStream(BinaryFile_Path+typeof(dicClass).Name+".CP3",FileMode.Open,FileAccess.Read))
        {
            byte[] bytes = new byte[fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();

            int index = 0;
            int count = BitConverter.ToInt32(bytes,index);//�����ܹ�count������
            index += 4;

            int keyLength = BitConverter.ToInt32(bytes, index);//key�ĳ���
            index += 4;

            string keyName = Encoding.UTF8.GetString(bytes,index, keyLength);//��ȡkey����
            index += keyLength;

            Type fieldType = typeof(FieldClass);
            FieldInfo[] fieldInfos = fieldType.GetFields();//�õ��ֶ����е��ֶ���Ϣ

            Type dicType = typeof(dicClass);
            object dicObj = Activator.CreateInstance(dicType);//ʵ�����ֵ������ 

            for (int i = 0; i < count; i++)
            {
                object fieldObj = Activator.CreateInstance(fieldType);
                foreach (FieldInfo field in fieldInfos)
                {
                    if (field.FieldType == typeof(int))
                    {
                        //�൱�ھ��ǰ�2��������תΪint Ȼ��ֵ���˶�Ӧ���ֶ�
                        field.SetValue(fieldObj, BitConverter.ToInt32(bytes, index));
                        index += 4;
                    }
                    else if (field.FieldType == typeof(float))
                    {
                        field.SetValue(fieldObj, BitConverter.ToSingle(bytes, index));
                        index += 4;
                    }
                    else if (field.FieldType == typeof(bool))
                    {
                        field.SetValue(fieldObj, BitConverter.ToBoolean(bytes, index));
                        index += 1;
                    }
                    else if (field.FieldType == typeof(string))
                    {
                        //��ȡ�ַ����ֽ�����ĳ���
                        int length = BitConverter.ToInt32(bytes, index);
                        index += 4;
                        field.SetValue(fieldObj, Encoding.UTF8.GetString(bytes, index, length));
                        index += length;
                    }                    
                }
                object dic = dicType.GetField("dataDic").GetValue(dicObj);
                MethodInfo method = dic.GetType().GetMethod("Add");
                object keyValue = fieldType.GetField(keyName).GetValue(fieldObj);
                method.Invoke(dic, new object[] { keyValue, fieldObj });
            }
            tableDic.Add(typeof(dicClass).Name,dicObj);
            fs.Close();
        }
    }
    /// <summary>
    /// �õ��������
    /// </summary>
    /// <typeparam name="T">����</typeparam>
    /// <returns></returns>
    public T GetTable<T>() where T:class
    {
       string tableName= typeof(T).Name;
       if(tableDic.ContainsKey(tableName))
       {
           return tableDic[tableName] as T;
       }
        return default(T);
    }
}
