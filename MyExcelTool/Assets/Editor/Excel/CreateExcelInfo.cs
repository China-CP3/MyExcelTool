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
    private static string BinaryFilePath="";//生成的二进制文件存放路径
    private static string fieldClassPath = Application.dataPath + "/Scripts/ExcelData/fieldClass/";//生成的字段文件存放路径
    private static string dicClassPath = Application.dataPath + "/Scripts/ExcelData/dicClass/";//生成的字典文件存放路径
    private static string ExcelFilePath = Application.dataPath+"/Editor/Excel/ExcelFile/";//excel文件存放路径
    /// <summary>
    /// 读取Excel表中的数据  生成3个文件
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
            using (FileStream fs = files[i].Open(FileMode.Open,FileAccess.Read))//读取第i个文件中所有表的数据 固定写法
            {
                IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                dataTable = excelDataReader.AsDataSet().Tables;
                fs.Close();
            }
            for (int j = 0; j < dataTable.Count; j++)//该文件中有Count张表,每张表生成对应的3个文件
            {
                GenateBinaryFile(dataTable[j]);
                GenateFieldFile(dataTable[j]);
                GenateDicFilePath(dataTable[j]);
            }
        }
    }/// <summary>
     /// 生成Excel表对应的2进制数据
     /// </summary>
     /// <param name="table"></param>
    private static void GenateBinaryFile(DataTable table)
    {

    }
    /// <summary>
    /// 生成Excel表对应的字段类
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
        AssetDatabase.Refresh();//刷新Projecet目录
    }
    /// <summary>
    /// 生成Excel表对应的字典类
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
