using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System;
using System.IO;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

public class ExcelImporter : AssetPostprocessor
{
	class ExcelAssetInfo
	{
		public Type AssetType { get; set; }
		public ExcelAssetAttribute Attribute { get; set; } 
	}

	static void OnPostprocessAllAssets (string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
	{
		bool imported = false;
		foreach(string path in importedAssets)
		{
			if(Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx") 
			{
				var excelName = Path.GetFileNameWithoutExtension(path);
				if(excelName.StartsWith("~$")) continue;

				var info = FindAssetTypeByExcelName(excelName);

				if(info == null)
				{
					Debug.LogWarning("ExcelAssetScript is not found. Select " + path + " and excute 'Create/ExcelAssetScript' from create menu.");
					continue;
				}

				ImportExcel(path, info);
				imported = true;
			}
		}

		if(imported) 
		{
			AssetDatabase.SaveAssets();
			AssetDatabase.Refresh();
		}
	}

	static ExcelAssetInfo FindAssetTypeByExcelName(string excelName)
	{
		foreach(var assembly in AppDomain.CurrentDomain.GetAssemblies())
		{
			var type = assembly.GetType(excelName);
			if(type == null) continue;

			var attributes = type.GetCustomAttributes(typeof(ExcelAssetAttribute), false);
			if(attributes.Length == 0) continue;

			var attribute = (ExcelAssetAttribute)attributes[0];

			return new ExcelAssetInfo()
			{
				AssetType = type,
				Attribute = attribute
			};
		}
		return null;
	}

	static UnityEngine.Object LoadOrCreateAsset(string assetPath, Type assetType)
	{
		Directory.CreateDirectory(Path.GetDirectoryName(assetPath));

		var asset = AssetDatabase.LoadAssetAtPath(assetPath, assetType);

		if (asset == null)
		{
			asset = ScriptableObject.CreateInstance(assetType.Name);
			AssetDatabase.CreateAsset((ScriptableObject)asset, assetPath);
			asset.hideFlags = HideFlags.NotEditable;
		}

		return asset;
	}

	static IWorkbook LoadBook(string excelPath)
	{
		using(FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
		{
			if (Path.GetExtension(excelPath) == ".xls") return new HSSFWorkbook(stream);
			else return new XSSFWorkbook(stream);
		}
	}

	static List<string> GetFieldNamesFromSheetHeader(ISheet sheet)
	{
		IRow headerRow = sheet.GetRow(0);

		var fieldNames = new List<string>();
		for (int i = 0; i < headerRow.LastCellNum; i++)
		{
			var cell = headerRow.GetCell(i);
			if(cell == null || cell.CellType == CellType.Blank) break;
			fieldNames.Add(cell.StringCellValue);
		}
		return fieldNames;
	}

	static object CellToFieldObject(ICell cell, FieldInfo fieldInfo, bool isFormulaEvalute = false)
	{
		var type = isFormulaEvalute ? cell.CachedFormulaResultType : cell.CellType;

		switch(type)
		{
			case CellType.String:
				if (fieldInfo.FieldType.IsEnum) return Enum.Parse(fieldInfo.FieldType, cell.StringCellValue);
				else return cell.StringCellValue;
			case CellType.Boolean:
				return cell.BooleanCellValue;
			case CellType.Numeric:
				return Convert.ChangeType(cell.NumericCellValue, fieldInfo.FieldType);
			case CellType.Formula:
				if(isFormulaEvalute) return null;
				return CellToFieldObject(cell, fieldInfo, true); 
			default:
				return null;
		}
	}

	static object CreateEntityFromRow(IRow row, List<string> columnNames, Type entityType)
	{
		var entity = Activator.CreateInstance(entityType);

		for (int i = 0; i < columnNames.Count; i++)
		{
			FieldInfo entityField = entityType.GetField(columnNames[i]);
			if (entityField == null) continue;

			ICell cell = row.GetCell(i);

			try
			{
				object fieldValue = CellToFieldObject(cell, entityField);
				entityField.SetValue(entity, fieldValue);
			}
			catch
			{
				throw new Exception(string.Format("Invalid excel cell type at row {0}, column {1}.", row.RowNum, cell.ColumnIndex));
			}
		}
		return entity;
	}

	static object GetEntityListFromSheet(ISheet sheet, Type entityType)
	{
		List<string> excelColumnNames = GetFieldNamesFromSheetHeader(sheet);

		Type listType = typeof(List<>).MakeGenericType(entityType);
		MethodInfo listAddMethod = listType.GetMethod("Add", new Type[]{entityType});
		object list = Activator.CreateInstance(listType);

		// row of index 0 is header
		for (int i = 1; i <= sheet.LastRowNum; i++)
		{
			IRow row = sheet.GetRow(i);
			if(row == null) break;

			ICell entryCell = row.GetCell(0); 
			if(entryCell == null || entryCell.CellType == CellType.Blank) break;

			// skip comment row
			if(entryCell.CellType == CellType.String && entryCell.StringCellValue.StartsWith("#")) continue;

			var entity = CreateEntityFromRow(row, excelColumnNames, entityType);
			listAddMethod.Invoke(list, new object[] { entity });
		}
		return list;
	}

	static void ImportExcel(string excelPath, ExcelAssetInfo info)
	{
		string assetPath = "";
		string assetName = info.AssetType.Name + ".asset";

		if(string.IsNullOrEmpty(info.Attribute.AssetPath))
		{
			string basePath = Path.GetDirectoryName(excelPath);
			assetPath = Path.Combine(basePath, assetName);
		}else{
			var path = Path.Combine("Assets", info.Attribute.AssetPath);
			assetPath = Path.Combine(path, assetName);
		}
		UnityEngine.Object asset = LoadOrCreateAsset(assetPath, info.AssetType);

		IWorkbook book = LoadBook(excelPath);

		var assetFields = info.AssetType.GetFields();

		foreach (var assetField in assetFields)
		{
			ISheet sheet =  book.GetSheet(assetField.Name);
			if(sheet == null) continue;

			Type fieldType = assetField.FieldType;
			if(! fieldType.IsGenericType || (fieldType.GetGenericTypeDefinition() != typeof(List<>))) continue;

			Type[] types = fieldType.GetGenericArguments();
			Type entityType = types[0];

			object entities = GetEntityListFromSheet(sheet, entityType);
			assetField.SetValue(asset, entities);
		}

		EditorUtility.SetDirty(asset);
	}
}
