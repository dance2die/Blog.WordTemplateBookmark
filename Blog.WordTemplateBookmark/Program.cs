using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

				// Find table.
				// calculate text length
				// increase the table size
			}
			catch (Exception ex)
			{
				document?.Close();
			}
			finally
			{
				word?.Quit(WdSaveOptions.wdSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat);
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
