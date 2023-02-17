using Excel;
using PlasticGui.WorkspaceWindow.Home;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using UnityEditor;
using UnityEngine;

public class CreateExcelInfo 
{
    private static string BinaryFilePath="";//���ɵĶ������ļ����·��
    private static string fieldClassPath = Application.dataPath + "/Scripts/ExcelData/fieldClass/";//���ɵ��ֶ��ļ����·��
    private static string dicClassPath = Application.dataPath + "/Scripts/ExcelData/dicClass/";//���ɵ��ֵ��ļ����·��
    private static string ExcelFilePath = Application.dataPath+"/Editor/Excel/ExcelFile/";//excel�ļ����·��
    /// <summary>
    /// ��ȡExcel���е�����  ����3���ļ�
    /// </summary>
    [MenuItem("Tool/ReadExcelInfo")]
    private static void ReadExcelInfo()
    {
        DirectoryInfo directory = Directory.CreateDirectory(ExcelFilePath);
        FileInfo[] files = directory.GetFiles();
        DataTableCollection dataTable;
        for (int i = 0; i < files.Length; i++)
        {
            if (files[i].Extension!=".xlsx"&& files[i].Extension != ".xls")
            {
                continue;
            }
            using (FileStream fs = files[i].Open(FileMode.Open,FileAccess.Read))//��ȡ��i���ļ������б������ �̶�д��
            {
                IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                dataTable = excelDataReader.AsDataSet().Tables;
                fs.Close();
            }
            for (int j = 0; j < dataTable.Count; j++)//���ļ�����Count�ű�,ÿ�ű����ɶ�Ӧ��3���ļ�
            {
                GenateBinaryFile(dataTable[j]);
                GenateFieldFile(dataTable[j]);
                GenateDicFilePath(dataTable[j]);
            }
        }
    }/// <summary>
     /// ����Excel���Ӧ��2��������
     /// </summary>
     /// <param name="table"></param>
    private static void GenateBinaryFile(DataTable table)
    {

    }
    /// <summary>
    /// ����Excel���Ӧ���ֶ���
    /// </summary>
    /// <param name="table"></param>
    private static void GenateFieldFile(DataTable table)
    {
        if(!Directory.Exists(fieldClassPath))
        {
           Directory.CreateDirectory(fieldClassPath);
        }
        string str = "public class " + table.TableName + "\n{\n";
        DataRow rowName = GetVariableNameRow(table);
        DataRow rowType = GetVariableTypeRow(table);
        for (int i = 0; i < table.Columns.Count; i++)
        {
            str += "  public " + rowType[i].ToString() + " " + rowName[i].ToString() + ";\n";
        }
        str += "}";
        File.WriteAllText(fieldClassPath+table.TableName+".cs",str);
        AssetDatabase.Refresh();//ˢ��ProjecetĿ¼
    }
    /// <summary>
    /// ����Excel���Ӧ���ֵ���
    /// </summary>
    /// <param name="table"></param>
    private static void GenateDicFilePath(DataTable table)
    {

    }
    private static DataRow GetVariableNameRow(DataTable table)
    {
       return  table.Rows[0];
    }
    private static DataRow GetVariableTypeRow(DataTable table)
    {
        return table.Rows[1];
    }
    private static int GetKeyIndex(DataTable table)
    {
        DataRow row = table.Rows[2];
        for (int i = 0; i < table.Columns.Count; i++)
        {
            if (row[i].ToString()=="key")
            {
                return i;
            }
        }
        return 0;
    }
}
