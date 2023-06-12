using System;
using System.IO;
using System.Diagnostics;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;

namespace PdfCombiner
{
	internal class Program
	{
		private const int DefaultOffset = 10;
		private const XGraphicsUnit DefaultOffsetUnit = XGraphicsUnit.Millimeter;
		private const int DefaultRotate = -90;

		private static void Main(string[] args)
		{
			Console.Title = nameof(PdfCombiner);
			Console.CursorVisible = false;

			if (!IsTrue(args.Length == 0, "Throw a book(-s) at me.")) 
			{ 
				foreach (string arg in args)
				{
					var clock = Stopwatch.StartNew();
					Prepare(arg);
					clock.Stop();
					Console.WriteLine($"Elapsed {clock.ElapsedMilliseconds} ms");
					Console.WriteLine();
				}
			}

			Console.WriteLine("Press Enter to exit...");
			Console.ReadLine();
		}

		private static void Prepare(string fileName)
		{
			if (IsTrue(!File.Exists(fileName), $"File '{fileName}' does not exists."))
			{
				return;
			}

			Console.WriteLine(fileName);

			var source = PdfReader.Open(fileName, PdfDocumentOpenMode.Import);
			if (IsTrue(source.PageCount == 0, "Empty document."))
			{
				return;
			}

			var offset = new XUnit(GetOffset(fileName), DefaultOffsetUnit);
			Console.WriteLine($"Offset: {offset}");
			CalculateParameters(source, offset, out var topPagePart, out var bottomPagePart);

			var destination = MakeNewAndCopyInfo(source);
			foreach (var sourcePage in source.Pages)
			{
				AddPagePart(destination, sourcePage, topPagePart);
				AddPagePart(destination, sourcePage, bottomPagePart);
			}

			Save(destination, fileName, removeNameAfterLastDot: offset != 0);
		}

		private static PdfDocument MakeNewAndCopyInfo(PdfDocument source)
		{
			var destination = new PdfDocument { PageLayout = PdfPageLayout.SinglePage };
			destination.Info.Creator = source.Info.Creator;
			destination.Info.CreationDate = source.Info.CreationDate;
			destination.Info.Title = source.Info.Title;
			destination.Info.Author = source.Info.Author;
			destination.Info.Subject = source.Info.Subject;
			return destination;
		}

		private static int GetOffset(string fileName)
		{
			fileName = Path.GetFileNameWithoutExtension(fileName);
			var extension = Path.GetExtension(fileName);
			
			if (!string.IsNullOrEmpty(extension) && int.TryParse(extension.Substring(1), out var offset))
			{
				return offset;
			}

			return DefaultOffset;
		}

		private static void CalculateParameters(PdfDocument source, XUnit offset, out PdfRectangle topPagePart, out PdfRectangle bottomPagePart)
		{
			var destinationPage = source.Pages[0];
			var halfPageHeight = destinationPage.Height / 2;
			var halfPageSize = new XSize(destinationPage.Width - (offset * 2), halfPageHeight - offset);
			topPagePart = new PdfRectangle(new XPoint(offset, halfPageHeight), halfPageSize);
			bottomPagePart = new PdfRectangle(new XPoint(offset, offset), halfPageSize);
		}

		private static void AddPagePart(PdfDocument destination, PdfPage sourcePage, PdfRectangle pagePart)
		{
			var destinationPage = destination.AddPage(sourcePage);
			destinationPage.CropBox = pagePart;
			destinationPage.Rotate = DefaultRotate;
		}

		private static void Save(PdfDocument destination, string originalFileName, bool removeNameAfterLastDot)
		{
			var path = Path.GetDirectoryName(originalFileName);
			var fileName = Path.GetFileNameWithoutExtension(originalFileName);
			if (removeNameAfterLastDot)
			{
				fileName = Path.GetFileNameWithoutExtension(fileName);
			}
			var extension = Path.GetExtension(originalFileName);
			var output = Path.Combine(path, $"{fileName}_prepared{extension}");
			destination.Save(output);

			Console.WriteLine($"Saved to '{output}'");

			Process.Start(output);
		}

		private static bool IsTrue(bool condition, string message)
		{
			if (condition)
			{
				var prevColor = Console.ForegroundColor;
				Console.ForegroundColor = ConsoleColor.Red;
				Console.WriteLine(message);
				Console.ForegroundColor = prevColor;
			}

			return condition;
		}
	}
}
