using System;
using System.Collections.Generic;
using System.Text;

namespace XlsPrinter.Core
{
    /// <summary>
    /// Represents a single piece of text extracted from a GDI EMF print record,
    /// with its logical/page coordinates, font properties, and decoded text.
    /// </summary>
    public class ExtractedTextElement
    {
        public int PageNumber { get; set; }
        public string Text { get; set; }
        public float X { get; set; }          // Left position in millimeters or logical page units
        public float Y { get; set; }          // Top position
        public float Width { get; set; }      // Width of text box
        public float Height { get; set; }     // Height of text box
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool IsBold { get; set; }
        public bool IsRtl { get; set; }
        public bool Selected { get; set; }

        public float Right { get { return X + Width; } }
        public float Bottom { get { return Y + Height; } }
        public float CenterY { get { return Y + (Height / 2.0f); } }
        public float CenterX { get { return X + (Width / 2.0f); } }

        public ExtractedTextElement()
        {
            Text = string.Empty;
            FontName = "Arial";
            FontSize = 10f;
            Selected = true;
        }

        public override string ToString()
        {
            return string.Format("[P{0} ({1:F1},{2:F1}) {3}x{4}] {5}", PageNumber, X, Y, Width, Height, Text);
        }
    }

    /// <summary>
    /// Represents a single extracted page from the print job.
    /// </summary>
    public class ExtractedPage
    {
        public int PageNumber { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public List<ExtractedTextElement> Elements { get; set; }
        public byte[] RawEmfBytes { get; set; }

        public ExtractedPage(int pageNumber)
        {
            PageNumber = pageNumber;
            Width = 210f;  // Default A4 width (mm) or logical units
            Height = 297f; // Default A4 height (mm)
            Elements = new List<ExtractedTextElement>();
        }
    }

    /// <summary>
    /// Represents an entire captured print job.
    /// </summary>
    public class PrintJobDocument
    {
        public int JobId { get; set; }
        public string DocumentTitle { get; set; }
        public string PrinterName { get; set; }
        public string UserName { get; set; }
        public DateTime CaptureTime { get; set; }
        public List<ExtractedPage> Pages { get; set; }

        public PrintJobDocument()
        {
            DocumentTitle = "Untitled Print Job";
            PrinterName = "XLS Virtual Printer";
            UserName = Environment.UserName;
            CaptureTime = DateTime.Now;
            Pages = new List<ExtractedPage>();
        }
    }
}
