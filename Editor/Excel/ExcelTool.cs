using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using Excel;
using UnityEditor;
using UnityEngine;

/// <summary>配置表代码与二进制文件导出工具。</summary>
public class ExcelTool
{
    /// <summary>Excel文件存放路径。</summary>
    public static readonly string EXCEL_PATH = Application.dataPath + "/ArtRes/Excel/";
    /// <summary>数据结构类脚本储存位置路径。</summary>
    public static readonly string DATA_CLASS_PATH = Application.dataPath + "/Scripts/ExcelData/DataClass/";
    /// <summary>容器类脚本储存位置路径。</summary>
    public static readonly string DATA_CONTAINER_PATH = Application.dataPath + "/Scripts/ExcelData/Container/";
    /// <summary>每个Excel文件中固定的读取规则配置表名称。</summary>
    private const string CONFIG_TABLE_NAME = "Config";
    /// <summary>Config表中工作表名所在的列索引。</summary>
    private const int CONFIG_SHEET_NAME_COLUMN = 0;
    /// <summary>Config表中数据起始行所在的列索引。</summary>
    private const int CONFIG_DATA_START_ROW_COLUMN = 1;
    /// <summary>Config表中数据起始列所在的列索引。</summary>
    private const int CONFIG_DATA_START_COLUMN = 2;
    /// <summary>Config表中变量名行所在的列索引。</summary>
    private const int CONFIG_FIELD_NAME_ROW_COLUMN = 3;
    /// <summary>Config表中变量类型行所在的列索引。</summary>
    private const int CONFIG_FIELD_TYPE_ROW_COLUMN = 4;
    /// <summary>Config表中主键列所在的列索引。</summary>
    private const int CONFIG_KEY_COLUMN = 5;
    /// <summary>Config表必须具备的固定列数量。</summary>
    private const int CONFIG_COLUMN_COUNT = 6;

    /// <summary>记录一张数据表经过校验后的读取范围。</summary>
    private sealed class TableExportInfo
    {
        /// <summary>来源Excel文件名。</summary>
        public string FileName;
        /// <summary>待导出的数据工作表。</summary>
        public DataTable Table;
        /// <summary>Config表中对应记录的Excel行号。</summary>
        public int ConfigRowNumber;
        /// <summary>数据起始行的零基索引。</summary>
        public int DataStartRowIndex;
        /// <summary>变量名行的零基索引。</summary>
        public int FieldNameRowIndex;
        /// <summary>变量类型行的零基索引。</summary>
        public int FieldTypeRowIndex;
        /// <summary>字段起始列的零基索引。</summary>
        public int FieldStartColumnIndex;
        /// <summary>字段数量。</summary>
        public int FieldCount;
        /// <summary>主键列的零基索引。</summary>
        public int KeyColumnIndex;
        /// <summary>数据结束行的零基排他索引。</summary>
        public int DataEndRowExclusive;
    }

    [MenuItem("GameTool/GenerateExcel")]
    private static void GenerateExcelInfo()
    {
        List<TableExportInfo> tables = new();
        List<string> errors = new();
        foreach (FileInfo file in Directory.CreateDirectory(EXCEL_PATH).GetFiles())
        {
            if (file.Extension != ".xlsx" && file.Extension != ".xls") continue;
            ReadWorkbook(file, tables, errors);
        }

        ValidateTables(tables, errors);
        if (errors.Count > 0)
        {
            string message = "配置表格式错误，已停止导出：\n" + string.Join("\n", errors);
            Debug.LogError(message);
            EditorUtility.DisplayDialog("配置表导出失败", message, "确定");
            return;
        }

        foreach (TableExportInfo table in tables)
        {
            GenerateExcelDataClass(table);
            GenerateExcelContainer(table);
            GenerateExcelBinary(table);
        }

        AssetDatabase.Refresh();
        string successMessage = $"导出完毕，共导出 {tables.Count} 个配置表。";
        Debug.Log(successMessage);
        EditorUtility.DisplayDialog("配置表导出", successMessage, "确定");
    }

    /// <summary>读取一个Excel文件，并根据其中的Config表收集待导出的数据表。</summary>
    /// <param name="file">Excel文件。</param>
    /// <param name="tables">待导出数据表列表。</param>
    /// <param name="errors">格式错误列表。</param>
    private static void ReadWorkbook(FileInfo file, List<TableExportInfo> tables, List<string> errors)
    {
        try
        {
            using FileStream stream = file.Open(FileMode.Open, FileAccess.Read);
            using IExcelDataReader reader = file.Extension == ".xls" ? ExcelReaderFactory.CreateBinaryReader(stream) : ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataTableCollection workbookTables = reader.AsDataSet().Tables;
            ReadWorkbookConfig(file.Name, workbookTables, tables, errors);
        }
        catch (Exception exception)
        {
            errors.Add(FormatError(file.Name, "未知", 0, 0, "读取Excel失败：" + exception.Message));
        }
    }

    /// <summary>读取并应用工作簿内固定名称的Config表。</summary>
    /// <param name="fileName">Excel文件名。</param>
    /// <param name="workbookTables">工作簿的全部工作表。</param>
    /// <param name="tables">待导出数据表列表。</param>
    /// <param name="errors">格式错误列表。</param>
    private static void ReadWorkbookConfig(string fileName, DataTableCollection workbookTables, List<TableExportInfo> tables, List<string> errors)
    {
        DataTable configTable = FindTable(workbookTables, CONFIG_TABLE_NAME);
        if (configTable == null)
        {
            errors.Add(FormatError(fileName, CONFIG_TABLE_NAME, 0, 0, "缺少固定名称的Config配置表。"));
            return;
        }
        if (configTable.Columns.Count < CONFIG_COLUMN_COUNT)
        {
            errors.Add(FormatError(fileName, CONFIG_TABLE_NAME, 0, 0, "Config表至少需要A到F共6列。"));
            return;
        }
        if (configTable.Rows.Count == 0)
        {
            errors.Add(FormatError(fileName, CONFIG_TABLE_NAME, 0, 0, "Config表至少需要第1行标题。"));
            return;
        }

        HashSet<string> configuredTableNames = new();
        for (int rowIndex = 1; rowIndex < configTable.Rows.Count; rowIndex++)
        {
            DataRow configRow = configTable.Rows[rowIndex];
            if (IsConfigRowEmpty(configRow)) continue;
            if (!TryCreateExportInfo(fileName, workbookTables, configRow, rowIndex, out TableExportInfo info, out string error))
            {
                errors.Add(error);
                continue;
            }
            if (!configuredTableNames.Add(info.Table.TableName))
            {
                errors.Add(FormatError(fileName, CONFIG_TABLE_NAME, rowIndex + 1, CONFIG_SHEET_NAME_COLUMN + 1, "工作表[" + info.Table.TableName + "]被重复配置。"));
                continue;
            }
            tables.Add(info);
        }

        foreach (DataTable table in workbookTables)
        {
            if (table.TableName == CONFIG_TABLE_NAME || table.Rows.Count == 0) continue;
            if (!configuredTableNames.Contains(table.TableName))
                errors.Add(FormatError(fileName, table.TableName, 0, 0, "非空工作表未在Config表中登记。"));
        }
    }

    /// <summary>从一行Config记录创建数据表导出信息。</summary>
    /// <param name="fileName">Excel文件名。</param>
    /// <param name="workbookTables">工作簿的全部工作表。</param>
    /// <param name="configRow">当前Config记录。</param>
    /// <param name="rowIndex">当前Config记录的零基行索引。</param>
    /// <param name="info">创建成功后的导出信息。</param>
    /// <param name="error">创建失败时的定位错误。</param>
    /// <returns>Config记录完整且基本合法时返回true。</returns>
    private static bool TryCreateExportInfo(string fileName, DataTableCollection workbookTables, DataRow configRow, int rowIndex, out TableExportInfo info, out string error)
    {
        info = null;
        error = null;
        int configRowNumber = rowIndex + 1;
        for (int column = 0; column < CONFIG_COLUMN_COUNT; column++)
        {
            if (IsCellEmpty(configRow[column]))
            {
                error = FormatError(fileName, CONFIG_TABLE_NAME, configRowNumber, column + 1, "Config记录不完整，A到F列均不能为空。" );
                return false;
            }
        }

        string tableName = configRow[CONFIG_SHEET_NAME_COLUMN].ToString();
        if (tableName == CONFIG_TABLE_NAME)
        {
            error = FormatError(fileName, CONFIG_TABLE_NAME, configRowNumber, CONFIG_SHEET_NAME_COLUMN + 1, "Config表本身不能作为数据表登记。" );
            return false;
        }
        DataTable table = FindTable(workbookTables, tableName);
        if (table == null)
        {
            error = FormatError(fileName, CONFIG_TABLE_NAME, configRowNumber, CONFIG_SHEET_NAME_COLUMN + 1, "找不到工作表[" + tableName + "]。" );
            return false;
        }

        if (!TryReadPositiveCoordinate(fileName, configRowNumber, CONFIG_DATA_START_ROW_COLUMN, configRow, "数据起始行", out int dataStartRow, out error) ||
            !TryReadPositiveCoordinate(fileName, configRowNumber, CONFIG_DATA_START_COLUMN, configRow, "数据起始列", out int dataStartColumn, out error) ||
            !TryReadPositiveCoordinate(fileName, configRowNumber, CONFIG_FIELD_NAME_ROW_COLUMN, configRow, "变量名行", out int fieldNameRow, out error) ||
            !TryReadPositiveCoordinate(fileName, configRowNumber, CONFIG_FIELD_TYPE_ROW_COLUMN, configRow, "变量类型行", out int fieldTypeRow, out error) ||
            !TryReadPositiveCoordinate(fileName, configRowNumber, CONFIG_KEY_COLUMN, configRow, "主键列", out int keyColumn, out error))
            return false;

        info = new TableExportInfo
        {
            FileName = fileName,
            Table = table,
            ConfigRowNumber = configRowNumber,
            DataStartRowIndex = dataStartRow - 1,
            FieldStartColumnIndex = dataStartColumn - 1,
            FieldNameRowIndex = fieldNameRow - 1,
            FieldTypeRowIndex = fieldTypeRow - 1,
            KeyColumnIndex = keyColumn - 1,
        };
        return true;
    }

    /// <summary>读取Config中的正整数坐标。</summary>
    /// <param name="fileName">Excel文件名。</param>
    /// <param name="configRowNumber">Config记录的Excel行号。</param>
    /// <param name="columnIndex">Config列的零基索引。</param>
    /// <param name="configRow">当前Config记录。</param>
    /// <param name="coordinateName">坐标名称。</param>
    /// <param name="value">读取成功后的坐标值。</param>
    /// <param name="error">读取失败时的定位错误。</param>
    /// <returns>单元格为正整数时返回true。</returns>
    private static bool TryReadPositiveCoordinate(string fileName, int configRowNumber, int columnIndex, DataRow configRow, string coordinateName, out int value, out string error)
    {
        if (int.TryParse(configRow[columnIndex].ToString(), out value) && value > 0)
        {
            error = null;
            return true;
        }
        error = FormatError(fileName, CONFIG_TABLE_NAME, configRowNumber, columnIndex + 1, coordinateName + "必须是大于0的整数。" );
        return false;
    }

    /// <summary>校验全部待导出数据表，确保写入前发现所有格式错误。</summary>
    /// <param name="tables">待导出的数据表。</param>
    /// <param name="errors">格式错误列表。</param>
    private static void ValidateTables(List<TableExportInfo> tables, List<string> errors)
    {
        HashSet<string> typeNames = new();
        foreach (TableExportInfo info in tables)
        {
            if (!IsValidIdentifier(info.Table.TableName))
                errors.Add(FormatError(info.FileName, info.Table.TableName, 0, 0, "工作表名不是合法的C#类型名。"));
            AddTypeName(info, info.Table.TableName, typeNames, errors);
            AddTypeName(info, info.Table.TableName + "Container", typeNames, errors);
            ValidateTable(info, errors);
        }
    }

    /// <summary>记录将生成的类型名，并检测命名冲突。</summary>
    /// <param name="info">数据表导出信息。</param>
    /// <param name="typeName">待生成类型名。</param>
    /// <param name="typeNames">已记录类型名集合。</param>
    /// <param name="errors">格式错误列表。</param>
    private static void AddTypeName(TableExportInfo info, string typeName, HashSet<string> typeNames, List<string> errors)
    {
        if (!typeNames.Add(typeName))
            errors.Add(FormatError(info.FileName, info.Table.TableName, 0, 0, "生成的类型名[" + typeName + "]与其他配置表冲突。"));
    }

    /// <summary>校验一张数据表的Config坐标、字段范围和数据内容。</summary>
    /// <param name="info">数据表导出信息。</param>
    /// <param name="errors">格式错误列表。</param>
    private static void ValidateTable(TableExportInfo info, List<string> errors)
    {
        DataTable table = info.Table;
        if (info.FieldNameRowIndex >= table.Rows.Count || info.FieldTypeRowIndex >= table.Rows.Count)
        {
            errors.Add(FormatError(info.FileName, CONFIG_TABLE_NAME, info.ConfigRowNumber, 0, "变量名行或变量类型行超出工作表[" + table.TableName + "]范围。"));
            return;
        }
        if (info.DataStartRowIndex > table.Rows.Count)
        {
            errors.Add(FormatError(info.FileName, CONFIG_TABLE_NAME, info.ConfigRowNumber, CONFIG_DATA_START_ROW_COLUMN + 1, "数据起始行最多只能是工作表末行的下一行。"));
            return;
        }
        if (info.FieldStartColumnIndex >= table.Columns.Count || info.KeyColumnIndex >= table.Columns.Count)
        {
            errors.Add(FormatError(info.FileName, CONFIG_TABLE_NAME, info.ConfigRowNumber, 0, "数据起始列或主键列超出工作表[" + table.TableName + "]范围。"));
            return;
        }
        if (info.DataStartRowIndex <= info.FieldNameRowIndex || info.DataStartRowIndex <= info.FieldTypeRowIndex)
        {
            errors.Add(FormatError(info.FileName, CONFIG_TABLE_NAME, info.ConfigRowNumber, CONFIG_DATA_START_ROW_COLUMN + 1, "数据起始行必须位于变量名行和变量类型行之后。"));
            return;
        }

        int fieldEnd = info.FieldStartColumnIndex;
        DataRow nameRow = table.Rows[info.FieldNameRowIndex];
        while (fieldEnd < table.Columns.Count && !IsCellEmpty(nameRow[fieldEnd])) fieldEnd++;
        info.FieldCount = fieldEnd - info.FieldStartColumnIndex;
        if (info.FieldCount == 0)
        {
            errors.Add(FormatError(info.FileName, table.TableName, info.FieldNameRowIndex + 1, info.FieldStartColumnIndex + 1, "数据起始列没有字段名。"));
            return;
        }
        if (info.KeyColumnIndex < info.FieldStartColumnIndex || info.KeyColumnIndex >= fieldEnd)
        {
            errors.Add(FormatError(info.FileName, CONFIG_TABLE_NAME, info.ConfigRowNumber, CONFIG_KEY_COLUMN + 1, "主键列必须位于字段读取范围内。"));
            return;
        }

        HashSet<string> fieldNames = new();
        DataRow typeRow = table.Rows[info.FieldTypeRowIndex];
        for (int column = info.FieldStartColumnIndex; column < fieldEnd; column++)
        {
            string fieldName = nameRow[column].ToString();
            string fieldType = typeRow[column].ToString();
            if (!IsValidIdentifier(fieldName))
                errors.Add(FormatError(info.FileName, table.TableName, info.FieldNameRowIndex + 1, column + 1, "字段名不是合法的C#标识符。"));
            else if (!fieldNames.Add(fieldName))
                errors.Add(FormatError(info.FileName, table.TableName, info.FieldNameRowIndex + 1, column + 1, "字段名[" + fieldName + "]重复。"));
            if (!IsSupportedType(fieldType))
                errors.Add(FormatError(info.FileName, table.TableName, info.FieldTypeRowIndex + 1, column + 1, "字段类型[" + fieldType + "]不受支持，只允许int、float、bool、string。"));
        }

        HashSet<object> keys = new();
        bool reachedDataEnd = false;
        info.DataEndRowExclusive = table.Rows.Count;
        for (int rowIndex = info.DataStartRowIndex; rowIndex < table.Rows.Count; rowIndex++)
        {
            DataRow row = table.Rows[rowIndex];
            if (IsCellEmpty(row[info.KeyColumnIndex]))
            {
                if (!reachedDataEnd)
                {
                    reachedDataEnd = true;
                    info.DataEndRowExclusive = rowIndex;
                }
                if (HasDataInFieldRange(row, info.FieldStartColumnIndex, info.FieldCount))
                    errors.Add(FormatError(info.FileName, table.TableName, rowIndex + 1, info.KeyColumnIndex + 1, "主键为空的行不能包含字段数据。"));
                continue;
            }
            if (reachedDataEnd)
            {
                errors.Add(FormatError(info.FileName, table.TableName, rowIndex + 1, info.KeyColumnIndex + 1, "主键列存在断档，空主键后不能继续出现数据。"));
                continue;
            }

            for (int column = info.FieldStartColumnIndex; column < fieldEnd; column++)
            {
                string type = typeRow[column].ToString();
                if (!TryParseCell(row[column].ToString(), type, out object value))
                {
                    errors.Add(FormatError(info.FileName, table.TableName, rowIndex + 1, column + 1, "值[" + row[column] + "]无法解析为" + type + "。"));
                    continue;
                }
                if (column == info.KeyColumnIndex && !keys.Add(value))
                    errors.Add(FormatError(info.FileName, table.TableName, rowIndex + 1, column + 1, "主键值[" + row[column] + "]重复。"));
            }
        }
    }

    /// <summary>判断Config记录的固定六列是否全部为空。</summary>
    /// <param name="row">Config记录。</param>
    /// <returns>固定六列全部为空时返回true。</returns>
    private static bool IsConfigRowEmpty(DataRow row)
    {
        for (int column = 0; column < CONFIG_COLUMN_COUNT; column++)
            if (!IsCellEmpty(row[column])) return false;
        return true;
    }

    /// <summary>判断一行中的字段读取范围是否包含任意非空单元格。</summary>
    /// <param name="row">待检查的数据行。</param>
    /// <param name="startColumnIndex">字段起始列的零基索引。</param>
    /// <param name="fieldCount">字段数量。</param>
    /// <returns>字段范围内存在非空单元格时返回true。</returns>
    private static bool HasDataInFieldRange(DataRow row, int startColumnIndex, int fieldCount)
    {
        for (int offset = 0; offset < fieldCount; offset++)
            if (!IsCellEmpty(row[startColumnIndex + offset])) return true;
        return false;
    }

    /// <summary>判断单元格是否为空。</summary>
    /// <param name="value">单元格值。</param>
    /// <returns>空值或空文本时返回true。</returns>
    private static bool IsCellEmpty(object value) => value == null || value == DBNull.Value || string.IsNullOrEmpty(value.ToString());

    /// <summary>按名称查找工作簿中的工作表。</summary>
    /// <param name="tables">工作簿的全部工作表。</param>
    /// <param name="tableName">工作表名称。</param>
    /// <returns>找到的工作表；不存在时返回空。</returns>
    private static DataTable FindTable(DataTableCollection tables, string tableName)
    {
        foreach (DataTable table in tables)
            if (table.TableName == tableName) return table;
        return null;
    }

    /// <summary>尝试按声明类型解析单元格内容。</summary>
    /// <param name="text">单元格文本。</param>
    /// <param name="type">字段类型。</param>
    /// <param name="value">解析成功后的值。</param>
    /// <returns>可解析时返回true。</returns>
    private static bool TryParseCell(string text, string type, out object value)
    {
        value = null;
        if (type == "string") { value = text; return true; }
        if (type == "int" && int.TryParse(text, out int intValue)) { value = intValue; return true; }
        if (type == "float" && float.TryParse(text, out float floatValue)) { value = floatValue; return true; }
        if (type == "bool" && bool.TryParse(text, out bool boolValue)) { value = boolValue; return true; }
        return false;
    }

    /// <summary>判断字段类型是否被当前二进制格式支持。</summary>
    /// <param name="type">字段类型文本。</param>
    /// <returns>受支持时返回true。</returns>
    private static bool IsSupportedType(string type) => type == "int" || type == "float" || type == "bool" || type == "string";

    /// <summary>判断文本是否为非关键字的合法C#标识符。</summary>
    /// <param name="text">待校验文本。</param>
    /// <returns>合法时返回true。</returns>
    private static bool IsValidIdentifier(string text)
    {
        if (string.IsNullOrEmpty(text) || IsCSharpKeyword(text) || (!char.IsLetter(text[0]) && text[0] != '_')) return false;
        for (int i = 1; i < text.Length; i++) if (!char.IsLetterOrDigit(text[i]) && text[i] != '_') return false;
        return true;
    }

    /// <summary>判断文本是否为C#保留关键字或上下文关键字。</summary>
    /// <param name="text">待校验文本。</param>
    /// <returns>是关键字时返回true。</returns>
    private static bool IsCSharpKeyword(string text)
    {
        switch (text)
        {
            case "abstract": case "as": case "base": case "bool": case "break": case "byte": case "case": case "catch": case "char": case "checked": case "class": case "const": case "continue": case "decimal": case "default": case "delegate": case "do": case "double": case "else": case "enum": case "event": case "explicit": case "extern": case "false": case "finally": case "fixed": case "float": case "for": case "foreach": case "goto": case "if": case "implicit": case "in": case "int": case "interface": case "internal": case "is": case "lock": case "long": case "namespace": case "new": case "null": case "object": case "operator": case "out": case "override": case "params": case "private": case "protected": case "public": case "readonly": case "ref": case "return": case "sbyte": case "sealed": case "short": case "sizeof": case "stackalloc": case "static": case "string": case "struct": case "switch": case "this": case "throw": case "true": case "try": case "typeof": case "uint": case "ulong": case "unchecked": case "unsafe": case "ushort": case "using": case "virtual": case "void": case "volatile": case "while": case "add": case "alias": case "ascending": case "async": case "await": case "by": case "descending": case "dynamic": case "equals": case "from": case "get": case "global": case "group": case "into": case "join": case "let": case "nameof": case "on": case "orderby": case "partial": case "remove": case "select": case "set": case "unmanaged": case "value": case "var": case "when": case "where": case "yield":
                return true;
            default: return false;
        }
    }

    /// <summary>生成统一格式的配置表错误信息。</summary>
    /// <param name="fileName">Excel文件名。</param>
    /// <param name="tableName">工作表名。</param>
    /// <param name="row">从1开始的行号，0表示无具体行。</param>
    /// <param name="column">从1开始的列号，0表示无具体列。</param>
    /// <param name="reason">错误原因。</param>
    /// <returns>错误文本。</returns>
    private static string FormatError(string fileName, string tableName, int row, int column, string reason) => $"文件[{fileName}]，工作表[{tableName}]，行[{row}]，列[{column}]：{reason}";

    /// <summary>生成数据结构类。</summary>
    /// <param name="info">已通过校验的数据表导出信息。</param>
    private static void GenerateExcelDataClass(TableExportInfo info)
    {
        if (!Directory.Exists(DATA_CLASS_PATH)) Directory.CreateDirectory(DATA_CLASS_PATH);
        StringBuilder builder = new("public class " + info.Table.TableName + "\n{\n");
        for (int offset = 0; offset < info.FieldCount; offset++)
        {
            int column = info.FieldStartColumnIndex + offset;
            builder.Append("    public ").Append(info.Table.Rows[info.FieldTypeRowIndex][column]).Append(' ').Append(info.Table.Rows[info.FieldNameRowIndex][column]).Append(";\n");
        }
        builder.Append('}');
        File.WriteAllText(DATA_CLASS_PATH + info.Table.TableName + ".cs", builder.ToString());
    }

    /// <summary>生成数据容器类。</summary>
    /// <param name="info">已通过校验的数据表导出信息。</param>
    private static void GenerateExcelContainer(TableExportInfo info)
    {
        if (!Directory.Exists(DATA_CONTAINER_PATH)) Directory.CreateDirectory(DATA_CONTAINER_PATH);
        string keyType = info.Table.Rows[info.FieldTypeRowIndex][info.KeyColumnIndex].ToString();
        string content = "using System.Collections.Generic;\npublic class " + info.Table.TableName + "Container\n{\n";
        content += "    public Dictionary<" + keyType + ", " + info.Table.TableName + "> dataDic = new();\n}";
        File.WriteAllText(DATA_CONTAINER_PATH + info.Table.TableName + "Container.cs", content);
    }

    /// <summary>生成配置表的二进制数据。</summary>
    /// <param name="info">已通过校验的数据表导出信息。</param>
    private static void GenerateExcelBinary(TableExportInfo info)
    {
        if (!Directory.Exists(BinaryDataManager.DATA_BINARY_PATH)) Directory.CreateDirectory(BinaryDataManager.DATA_BINARY_PATH);
        using FileStream stream = new(BinaryDataManager.DATA_BINARY_PATH + info.Table.TableName + ".wang", FileMode.Create, FileAccess.Write);
        stream.Write(BitConverter.GetBytes(info.DataEndRowExclusive - info.DataStartRowIndex), 0, 4);
        byte[] bytes = Encoding.UTF8.GetBytes(info.Table.Rows[info.FieldNameRowIndex][info.KeyColumnIndex].ToString());
        stream.Write(BitConverter.GetBytes(bytes.Length), 0, 4);
        stream.Write(bytes, 0, bytes.Length);

        for (int rowIndex = info.DataStartRowIndex; rowIndex < info.DataEndRowExclusive; rowIndex++)
            for (int offset = 0; offset < info.FieldCount; offset++)
            {
                int column = info.FieldStartColumnIndex + offset;
                string value = info.Table.Rows[rowIndex][column].ToString();
                switch (info.Table.Rows[info.FieldTypeRowIndex][column].ToString())
                {
                    case "int": bytes = BitConverter.GetBytes(int.Parse(value)); stream.Write(bytes, 0, 4); break;
                    case "float": bytes = BitConverter.GetBytes(float.Parse(value)); stream.Write(bytes, 0, 4); break;
                    case "bool": bytes = BitConverter.GetBytes(bool.Parse(value)); stream.Write(bytes, 0, 1); break;
                    case "string": bytes = Encoding.UTF8.GetBytes(value); stream.Write(BitConverter.GetBytes(bytes.Length), 0, 4); stream.Write(bytes, 0, bytes.Length); break;
                }
            }
    }
}
