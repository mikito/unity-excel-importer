Unity Excel Importer
====================

Extension for auto import data from xls, xlsx to custom ScriptableObject in Unity Editor.


### Package Manager

Add the following entry to Packages/manifest.json or Package Manager.

```json
{
  "dependencies": {
    "net.mikinya.unity-excel-importer": "https://github.com/mikito/unity-excel-importer.git?path=Assets/ExcelImporter#v0.1.1/upm",
    ...
  }
}
```

Import Setup
--------

### 1. Create Excel and add to Unity 

Create an Excel file, make the first row the name of the column, and enter the data from the second row. And add it to Unity's Project view.

<img width="229" alt="create_excel" src="https://user-images.githubusercontent.com/1071168/34470078-7246b924-ef6d-11e7-8711-acb49ae2c349.png">

<img width="247" alt="import_excel_to_unity" src="https://user-images.githubusercontent.com/1071168/34470091-8bdc1c76-ef6d-11e7-88e5-76c76294a26c.png">

### 2. Create Entity Class Script

Create a new script and define a class with Excel column names and public fields of the desired type. Also give the class 'System.Serializable' attribute.

```cs
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[System.Serializable]
public class MstItemEntity
{
	public int id;
	public string name;
	public int price; 
}

```

### 3. Create Excel Asset Script 

After selecting Excel, execute ExcelAssetScript from the Create menu and create a ScriptableObject script for Excel.

<img width="527" alt="create_excel_asset" src="https://user-images.githubusercontent.com/1071168/34470096-abdadd0a-ef6d-11e7-9bf6-e9458c660264.png">

As for the generated script, the Excel file name and the sheet name are extracted and the part is commented out as below.

```cs
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[ExcelAsset]
public class MstItems : ScriptableObject
{
	//public List<EntityType> Entities; // Replace 'EntityType' to an actual type that is serializable.
}
```

### 4. Replace EntityType in created Excel Asset

Uncomment fields and replace the generic type of List with the Entity class defined above.

```cs
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[ExcelAsset]
public class MstItems : ScriptableObject
{
	public List<MstItemEntity> Entities;
}
```

### 4. Reimport or re-save Excel

When you import or re-save Excel, a ScriptableObject with the same name as Excel is created in the same directory and the contents of Excel are imported.

<img width="321" alt="reimport_excel" src="https://user-images.githubusercontent.com/1071168/34470099-c21db84e-ef6d-11e7-954f-cc5c96c382bd.png">

<img width="570" alt="imported_entities" src="https://user-images.githubusercontent.com/1071168/34470102-cde2bcc4-ef6d-11e7-8083-a661b82ae707.png">

After this setting, updating Excel automatically updates ScriptableObject as well.

Advanced
----------------

### Comment Row

If you enter '#' in the first cell of the row, you can treat it as a comment and skip.

### Change Asset Path 

You can change the ScriptableObject generation position by specifying AssetPath as the ExcelAssetAttribute as shown below.

```cs
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[ExcelAsset(AssetPath = "Resources/MasterData")]
public class MstItems : ScriptableObject
{
	public List<MstItemEntity> Entities;
}
```

### Use Enum

You can use enum by entering the element name as string in cell.
It is also useful to set Data Validation pull down as an element of enum in Excel.

### Log On Import

When true is specified for LogOnImport of ExcelAssetAttribute, a log is output when the import process runs.

```cs
...
[ExcelAsset(LogOnImport = true)]
public class MstItems : ScriptableObject
{
...
```

### Changing name association between ExcelAsset and ExcelFile

You can change the association to a specific Excel file by specifying ExcelName of ExcelAssetAttribute, 

```cs
...
[ExcelAsset(ExcelName = "MstItemData")]
public class MstItems : ScriptableObject
{
...
```

License
----------

This library is under the MIT License.
This software includes the work that is distributed in the Apache License 2.0.
