using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System;
using System.Linq;

[Serializable]
public class MstItemEntity
{
	public int id;
	public string name;
	public int price;
	public bool isNotForSale;
	public float rate;
	public MstItemCategory category;
	public int[] subIds;
	public string subIdsString
	{
		set	
		{
			 subIds = value.Split(',').Select(s => int.Parse(s)).ToArray();
		}
	}
}

public enum MstItemCategory
{
	Red,
	Green,
	Blue,
}