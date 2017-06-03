using System;
using System.Drawing;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Blog.WordTemplateBookmark
{
	public class Program
	{
		public static void Main(string[] args)
		{
			Application word = null;
			Document document = null;

			try
			{
				// Open Document
				word = new Application();

				var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "BookmarkTest.docx");
				document = word.Documents.Open(templatePath);

				// Populate bookmarks.
				const string signer = "Anna Bertha Cecilia Diana Emily Fanny Gertrude Hypatia Inez Jane ";
				SetSignerBookmarks(word, document.Bookmarks, signer);

				// Find table and get table width (in Word, table collection is 1-based)
				var table = document.Tables[1];
				double tableWidth = GetTableWidth(table);

				// calculate text width
				var textWidth = CalculateTextWidth(signer);

				// increase the table size
				if (textWidth > tableWidth)
					UpdateTableWidth(table, textWidth);
			}
			finally
			{
				document?.Close();
				word?.Quit(WdSaveOptions.wdSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat);
			}
		}

		private static void UpdateTableWidth(Table table, double textWidth)
		{
			table.Range.Cells.Width = (float)textWidth;
		}

		private static double GetTableWidth(Table table)
		{
			return table.Range.Cells.Width;
		}

		private static double CalculateTextWidth(string text)
		{
			// https://www.aspose.com/community/forums/thread/332975/setting-cell-width-according-to-text-size-in-aspose-word-java-apis.aspx
			using (Bitmap bmp = new Bitmap(1, 1))
			{
				bmp.SetResolution(96, 96);
				using (Graphics g = Graphics.FromImage(bmp))
				{
					var familyName = new FontFamily("Calibri");
					const float fontSize = 10;
					using (System.Drawing.Font font = new System.Drawing.Font(familyName, fontSize))
					{
						return g.MeasureString(text, font).Width;
					}
				}
			}
		}

		private static void SetSignerBookmarks(Application word, Bookmarks bookmarks, string signer)
		{
			// https://social.msdn.microsoft.com/Forums/vstudio/en-US/32b25cfd-cc5b-4e9f-bcbf-0dbbd49bca02/how-to-replace-text-or-insert-text-into-bookmarks-in-a-word-template?forum=csharpgeneral
			bookmarks["signer1"].Select();
			word.Selection.TypeText(signer);

			bookmarks["signer2"].Select();
			word.Selection.TypeText(signer);
		}
	}
}
