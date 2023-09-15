﻿namespace LB1
{
	public static class CSVFileWriter
	{
		public static void SaveFile(string path, List<string> lines)
		{
			using StreamWriter sw = new(path, false, System.Text.Encoding.Unicode);
			foreach (var line in lines)
			{
				sw.WriteLine(line);
			}
		}
	}
}
