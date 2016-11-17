using System;
using System.IO;
using System.Reflection;
using System.Linq;
using OfficeOpenXml;

namespace ExcelGenelator
{
	class MainClass
	{
		public static void Main(string[] args)
		{
			//args = new String[] { "input.csv", "output.xlsx" };

			// アプリケーションのフルパス
			var appPath = Assembly.GetExecutingAssembly().Location;
			if (args.Length != 2)
			{
				Console.WriteLine($"use: mono {Path.GetFileName(appPath)} input.csv output.xlsx");
				return;
			}
			var appDirectory = Path.GetDirectoryName(appPath);

			// テンプレートExcel
			var templateExcelName = Path.Combine(appDirectory, "template.xlsx");
			if (!File.Exists(templateExcelName))
			{
				Console.WriteLine($"ERROR {templateExcelName} not Found.");
				return;
			}

			// 入力CSV
			var inputCsvName = Path.Combine(appDirectory, args[0]);
			if (!File.Exists(inputCsvName))
			{
				Console.WriteLine($"ERROR {inputCsvName} not Found.");
				return;
			}

			// 出力Excel
			var outputExcelName = Path.Combine(appDirectory, args[1]);
			var wb = new ExcelPackage(new FileInfo(templateExcelName));
			var sheet = wb.Workbook.Worksheets.First();
			var lines = File.ReadAllLines(inputCsvName);
			foreach (var item in lines.Select((line, row) => new { line, row }))
			{
				var values = item.line.Split(',');

				foreach (int i in Enumerable.Range(0, 3))
				{
					// データを差し込むのは7行目のカラム３個目以降(セル番号は、0からではなく1から指定)
					var cell = sheet.Cells[item.row + 8, i + 4];
					if (i == 0) // 品名は、文字として挿入
					{
						cell.Value = values[i];
					}
					else
					{ // 数量・単価は、数値として挿入
						cell.Value = Int32.Parse(values[i]);
					}
				}
				//数式の再計算は必要ありません。
				wb.SaveAs(new FileInfo(outputExcelName));
			}
		}
	}
}
