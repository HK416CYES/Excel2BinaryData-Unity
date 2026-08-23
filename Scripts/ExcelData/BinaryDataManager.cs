using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using UnityEngine;

public class BinaryDataManager
{
    private static BinaryDataManager instance = new();
    public static BinaryDataManager Instance => instance;

    /// <summary>二进制数据存储位置路径。</summary>
    public static readonly string DATA_BINARY_PATH = Application.streamingAssetsPath + "/Binary/";
    /// <summary>存储所有Excel表数据的容器。</summary>
    private Dictionary<string, object> tableDic = new();

    /// <summary>序列化对象的存储路径。</summary>
    private static readonly string SAVE_PATH = Application.persistentDataPath + "/Data/";
    /// <summary>序列化对象使用的文件扩展名。</summary>
    private const string EXTENSION_NAME = ".wang";

    private BinaryDataManager() { }

    /// <summary>清空旧表并初始化全部配置表数据。</summary>
    public void InitData()
    {
        tableDic.Clear();
        LoadTable<TowerInfoContainer, TowerInfo>();
        LoadTable<PlayerInfoContainer, PlayerInfo>();
        LoadTable<TestInfoContainer, TestInfo>();
    }

    /// <summary>加载Excel表的二进制数据到内存中。</summary>
    /// <typeparam name="T">容器类名。</typeparam>
    /// <typeparam name="K">数据结构类名。</typeparam>
    public void LoadTable<T, K>()
    {
        string tableName = typeof(K).Name;
        string filePath = DATA_BINARY_PATH + tableName + ".wang";
        try
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException("找不到二进制文件。", filePath);
            byte[] bytes = File.ReadAllBytes(filePath);
            int index = 0;
            int count = ReadInt32(bytes, ref index, "表头的行数");
            if (count < 0) throw new InvalidDataException("表头的行数不能为负数。");
            int keyNameLength = ReadInt32(bytes, ref index, "表头的主键名长度");
            string keyName = ReadString(bytes, ref index, keyNameLength, "表头的主键名");
            if (string.IsNullOrEmpty(keyName)) throw new InvalidDataException("表头的主键名不能为空。");

            Type containerType = typeof(T);
            Type classType = typeof(K);
            FieldInfo dictionaryField = containerType.GetField("dataDic");
            FieldInfo keyField = classType.GetField(keyName);
            if (dictionaryField == null) throw new InvalidDataException("容器类缺少dataDic字段。");
            if (keyField == null) throw new InvalidDataException("数据类缺少主键字段[" + keyName + "]。");

            object containerObj = Activator.CreateInstance(containerType);
            object dictionaryObj = dictionaryField.GetValue(containerObj);
            MethodInfo addMethod = dictionaryObj?.GetType().GetMethod("Add");
            if (addMethod == null) throw new InvalidDataException("容器的dataDic字段不支持Add方法。");

            FieldInfo[] fields = classType.GetFields();
            HashSet<object> keys = new();
            for (int rowIndex = 0; rowIndex < count; rowIndex++)
            {
                object dataObj = Activator.CreateInstance(classType);
                foreach (FieldInfo field in fields) ReadFieldValue(bytes, ref index, dataObj, field, rowIndex);
                object keyValue = keyField.GetValue(dataObj);
                if (!keys.Add(keyValue)) throw new InvalidDataException("第" + (rowIndex + 1) + "条数据的主键[" + keyValue + "]重复。");
                addMethod.Invoke(dictionaryObj, new[] { keyValue, dataObj });
            }
            if (index != bytes.Length) throw new InvalidDataException("数据读取结束后仍剩余" + (bytes.Length - index) + "个未消费字节，文件格式可能不匹配。");
            tableDic[typeof(T).Name] = containerObj;
        }
        catch (Exception exception)
        {
            Debug.LogError("加载二进制表失败：表[" + tableName + "]，文件[" + filePath + "]，原因：" + exception.Message);
        }
    }

    /// <summary>按字段类型读取并设置一条数据的字段值。</summary>
    /// <param name="bytes">完整二进制数据。</param>
    /// <param name="index">当前读取位置，会在读取后前移。</param>
    /// <param name="dataObj">当前数据对象。</param>
    /// <param name="field">待读取字段。</param>
    /// <param name="rowIndex">从0开始的数据行索引。</param>
    private static void ReadFieldValue(byte[] bytes, ref int index, object dataObj, FieldInfo field, int rowIndex)
    {
        string context = "第" + (rowIndex + 1) + "条数据的字段[" + field.Name + "]";
        if (field.FieldType == typeof(int)) field.SetValue(dataObj, ReadInt32(bytes, ref index, context));
        else if (field.FieldType == typeof(float)) field.SetValue(dataObj, ReadSingle(bytes, ref index, context));
        else if (field.FieldType == typeof(bool)) field.SetValue(dataObj, ReadBoolean(bytes, ref index, context));
        else if (field.FieldType == typeof(string))
        {
            int length = ReadInt32(bytes, ref index, context + "的字符串长度");
            field.SetValue(dataObj, ReadString(bytes, ref index, length, context));
        }
        else throw new InvalidDataException(context + "的类型[" + field.FieldType.Name + "]不受支持。");
    }

    /// <summary>从当前读取位置安全读取32位整数。</summary>
    /// <param name="bytes">完整二进制数据。</param>
    /// <param name="index">当前读取位置，会在读取后前移。</param>
    /// <param name="context">读取内容说明。</param>
    /// <returns>读取到的整数。</returns>
    private static int ReadInt32(byte[] bytes, ref int index, string context)
    {
        EnsureReadable(bytes, index, 4, context);
        int value = BitConverter.ToInt32(bytes, index);
        index += 4;
        return value;
    }

    /// <summary>从当前读取位置安全读取单精度浮点数。</summary>
    /// <param name="bytes">完整二进制数据。</param>
    /// <param name="index">当前读取位置，会在读取后前移。</param>
    /// <param name="context">读取内容说明。</param>
    /// <returns>读取到的浮点数。</returns>
    private static float ReadSingle(byte[] bytes, ref int index, string context)
    {
        EnsureReadable(bytes, index, 4, context);
        float value = BitConverter.ToSingle(bytes, index);
        index += 4;
        return value;
    }

    /// <summary>从当前读取位置安全读取布尔值。</summary>
    /// <param name="bytes">完整二进制数据。</param>
    /// <param name="index">当前读取位置，会在读取后前移。</param>
    /// <param name="context">读取内容说明。</param>
    /// <returns>读取到的布尔值。</returns>
    private static bool ReadBoolean(byte[] bytes, ref int index, string context)
    {
        EnsureReadable(bytes, index, 1, context);
        bool value = BitConverter.ToBoolean(bytes, index);
        index++;
        return value;
    }

    /// <summary>从当前读取位置安全读取UTF-8字符串。</summary>
    /// <param name="bytes">完整二进制数据。</param>
    /// <param name="index">当前读取位置，会在读取后前移。</param>
    /// <param name="length">字符串字节长度。</param>
    /// <param name="context">读取内容说明。</param>
    /// <returns>读取到的字符串。</returns>
    private static string ReadString(byte[] bytes, ref int index, int length, string context)
    {
        if (length < 0) throw new InvalidDataException(context + "的字符串长度不能为负数。");
        EnsureReadable(bytes, index, length, context);
        string value = Encoding.UTF8.GetString(bytes, index, length);
        index += length;
        return value;
    }

    /// <summary>确认即将读取的字节范围完整位于缓冲区内。</summary>
    /// <param name="bytes">完整二进制数据。</param>
    /// <param name="index">当前读取位置。</param>
    /// <param name="count">需要读取的字节数量。</param>
    /// <param name="context">读取内容说明。</param>
    private static void EnsureReadable(byte[] bytes, int index, int count, string context)
    {
        if (index < 0 || count < 0 || index > bytes.Length || count > bytes.Length - index)
            throw new InvalidDataException(context + "在字节偏移" + index + "处需要读取" + count + "字节，但文件总长度为" + bytes.Length + "。");
    }

    /// <summary>获取已加载的数据表容器。</summary>
    /// <typeparam name="T">容器类型。</typeparam>
    /// <returns>已加载的容器；未加载时返回空。</returns>
    public T GetTable<T>() where T : class
    {
        string tableName = typeof(T).Name;
        if (tableDic.ContainsKey(tableName)) return tableDic[tableName] as T;
        return null;
    }

    /// <summary>存储类对象为二进制数据。</summary>
    /// <param name="obj">要储存的对象。</param>
    /// <param name="fileName">文件名。</param>
    public void Save(object obj, string fileName)
    {
        if (!Directory.Exists(SAVE_PATH)) Directory.CreateDirectory(SAVE_PATH);
        using (FileStream fs = new(SAVE_PATH + fileName + EXTENSION_NAME, FileMode.OpenOrCreate, FileAccess.Write))
        {
            BinaryFormatter bf = new();
            bf.Serialize(fs, obj);
        }
    }

    /// <summary>读取二进制数据并反序列化为对象。</summary>
    /// <typeparam name="T">目标对象类型。</typeparam>
    /// <param name="fileName">文件名。</param>
    /// <returns>读取成功后的对象；文件不存在时返回空。</returns>
    public T Load<T>(string fileName) where T : class
    {
        if (!File.Exists(SAVE_PATH + fileName + EXTENSION_NAME))
        {
            Debug.LogWarning("该文件不存在：" + SAVE_PATH + fileName + EXTENSION_NAME);
            return default;
        }
        using FileStream fs = File.Open(SAVE_PATH + fileName + EXTENSION_NAME, FileMode.Open, FileAccess.Read);
        BinaryFormatter bf = new();
        return bf.Deserialize(fs) as T;
    }
}
