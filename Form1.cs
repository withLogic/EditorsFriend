using FileFormat.Words;
using FileFormat.Words.Properties;
using FileFormat.Words.Table;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Windows.Forms;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml.Drawing.Charts;
using Spire.Pdf;
using Microsoft.Office.Interop.Word;

namespace DocxEasyFormat
{
    public partial class EditorsFriend : Form
    {
        public EditorsFriend()
        {
            InitializeComponent();
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                listView1.Items.Clear();

                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (var file in files)
                {
                    /* 
                     *  This is where I would need to update the view to add the items and create progress bars for them
                     */

                    String fileName = Path.GetFileName(file);

                    ListViewItem tlvi = new ListViewItem();
                    tlvi.Text = fileName;
                    tlvi.SubItems.Add("New");

                    listView1.Items.Add(tlvi);

                    string fileType = Path.GetExtension(file).ToUpper();

                    switch(fileType)
                        {
                        case ".PDF":
                            Console.WriteLine("Converting PDF file.");
                            tlvi.SubItems[1].Text = "Processing";

                            try
                            {
                                ProcessPdfFile(file);
                                tlvi.SubItems[1].Text = "Complete";
                            } catch (Exception ex)
                            {
                                tlvi.SubItems[1].Text = "Error";
                            }

                            break;

                        case ".DOC":
                            Console.WriteLine("Converting DOC file.");
                            tlvi.SubItems[1].Text = "Processing";

                            try
                            {
                                ProcessOldDocFile(file);
                                tlvi.SubItems[1].Text = "Complete";
                            } catch (Exception ex) 
                            {
                                tlvi.SubItems[1].Text = "Error";
                            }

                            break;

                        case ".DOCX":
                            Console.WriteLine("Converting DOCX file.");
                            tlvi.SubItems[1].Text = "Processing";

                            try
                            {
                                ProcessWordFile(file);
                                tlvi.SubItems[1].Text = "Complete";
                            } catch (Exception ex)
                            {
                                tlvi.SubItems[1].Text = "Error";
                            }

                            break;

                        default:
                            tlvi.SubItems[1].Text = "Error";
                            break;
                        }

                }
            }
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void ProcessOldDocFile(string filePath)
        {
            string dir = Path.GetDirectoryName(filePath);
            string fn = Path.GetFileNameWithoutExtension(filePath);
            var fileNameString = dir + "\\" + fn + "_convertedFromPDF.docx";

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            // try and use Word to manipulate the PDF first.
            if (wordApp != null)
            {
                wordApp.Visible = false;
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();
                doc = wordApp.Documents.Open(filePath);
                doc.SaveAs2(fileNameString, WdSaveFormat.wdFormatXMLDocument);

                doc.Close();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);
            }
            // fall back on the freeSpire implemetnation
            else
            {
                PdfDocument pdfFile = new PdfDocument();
                pdfFile.LoadFromFile(filePath);

                if (pdfFile != null)
                {
                    pdfFile.ConvertOptions.SetPdfToDocOptions(true, true);
                    pdfFile.SaveToFile(fileNameString, Spire.Pdf.FileFormat.DOCX);
                    pdfFile.Close();
                }
            }

            if (File.Exists(fileNameString))
            {
                ProcessWordFile(fileNameString);
            }
        }

        private void ProcessPdfFile(string filePath)
        {
            string dir = Path.GetDirectoryName(filePath);
            string fn = Path.GetFileNameWithoutExtension(filePath);
            var fileNameString = dir + "\\" + fn + "_convertedFromPDF.docx";

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            // try and use Word to manipulate the PDF first.
            if (wordApp != null)
            {
                wordApp.Visible = false;
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();
                doc = wordApp.Documents.Open(filePath);
                doc.SaveAs2(fileNameString, WdSaveFormat.wdFormatXMLDocument);

                doc.Close();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);
            }
            // fall back on the freeSpire implemetnation
            else
            {
                PdfDocument pdfFile = new PdfDocument();
                pdfFile.LoadFromFile(filePath);

                if (pdfFile != null)
                {
                    pdfFile.ConvertOptions.SetPdfToDocOptions(true, true);
                    pdfFile.SaveToFile(fileNameString, Spire.Pdf.FileFormat.DOCX);
                    pdfFile.Close();
                }
            }

            if (File.Exists(fileNameString))
            {
                ProcessWordFile(fileNameString);
            }
        }

        private void ProcessWordFile(string filePath)
        {
            string dir = Path.GetDirectoryName(filePath);
            string fn = Path.GetFileNameWithoutExtension(filePath);

            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                MemoryStream ms = new MemoryStream();
                fs.CopyTo(ms);
                WordprocessingDocument doc = WordprocessingDocument.Open(ms, true);
                if (doc != null)
                {
                    if (doc.MainDocumentPart != null)
                    {
                        // delete the header and footer (and their references)
                        if (doc.MainDocumentPart.HeaderParts.Count() > 0 || doc.MainDocumentPart.FooterParts.Count() > 0)
                        {
                            doc.MainDocumentPart.DeleteParts(doc.MainDocumentPart.HeaderParts);
                            doc.MainDocumentPart.DeleteParts(doc.MainDocumentPart.FooterParts);

                            DocumentFormat.OpenXml.Wordprocessing.Document tDoc = doc.MainDocumentPart.Document;
                            var headerReferences = tDoc.Descendants<HeaderReference>().ToList();

                            foreach (var header in headerReferences)
                            {
                                header.Remove();
                            }

                            var footerReferences = tDoc.Descendants<FooterReference>().ToList();

                            foreach (var footer in footerReferences)
                            {
                                footer.Remove();
                            }
                        }

                        DocumentFormat.OpenXml.Wordprocessing.Body bod = doc.MainDocumentPart.Document.Body;

                        var paragraphs = bod.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

                        foreach (DocumentFormat.OpenXml.Wordprocessing.Paragraph para in paragraphs.ToArray())
                        {
                            // go through the paragraphs
                            if (para != null)
                            {
                                if (para.Elements<ParagraphProperties>().Count() == 0)
                                {
                                    para.PrependChild<ParagraphProperties>(new ParagraphProperties());
                                }

                                ParagraphProperties pPr = para.Elements<ParagraphProperties>().First();

                                // get the paragraph properties
                                if (pPr != null)
                                {
                                    // set the line spacing
                                    if (pPr.Elements<SpacingBetweenLines>().Count() == 0)
                                    {
                                        pPr.AppendChild<SpacingBetweenLines>(new SpacingBetweenLines() { Before = "0", After = "0", Line = "240" });
                                    }

                                    SpacingBetweenLines pSp = pPr.Elements<SpacingBetweenLines>().First();
                                    pSp.Before = "0";
                                    pSp.After = "0";
                                    pSp.Line = "240";

                                    if (pPr.Elements<ContextualSpacing>().Count() == 0)
                                    {
                                        pPr.AppendChild<ContextualSpacing>(new ContextualSpacing());
                                    }

                                    // remove any paragrah style
                                    if (pPr.Elements<ParagraphStyleId>().Count() > 0)
                                    {
                                        pPr.RemoveChild<ParagraphStyleId>(pPr.Elements<ParagraphStyleId>().First());
                                    }

                                    // remove any paragrah alignemnt horizontal alignment
                                    if (pPr.Elements<Justification>().Count() > 0)
                                    {
                                        pPr.RemoveChild<Justification>(pPr.Elements<Justification>().First());
                                    }

                                    // remove any paragrah alignemnt vertical alignment
                                    if (pPr.Elements<TextAlignment>().Count() > 0)
                                    {
                                        pPr.RemoveChild<TextAlignment>(pPr.Elements<TextAlignment>().First());
                                    }

                                    // set the paragraph mark style
                                    if (pPr.Elements<ParagraphMarkRunProperties>().Count() == 0)
                                    {
                                        pPr.PrependChild<ParagraphMarkRunProperties>(new ParagraphMarkRunProperties());
                                    }

                                    ParagraphMarkRunProperties rPr = pPr.Elements<ParagraphMarkRunProperties>().First();
                                    rPr.RemoveAllChildren();

                                    RunFonts runFont = new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
                                    DocumentFormat.OpenXml.Wordprocessing.Color runColor = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "000000" };
                                    FontSize runSize = new FontSize { Val = new StringValue("22") };
                                    FontSizeComplexScript runSizeCS = new FontSizeComplexScript { Val = new StringValue("22") };

                                    rPr.AppendChild(runFont);
                                    rPr.AppendChild(runColor);
                                    rPr.AppendChild(runSize);
                                    rPr.AppendChild(runSizeCS);

                                    // remove any page breaks
                                    if (para.Elements<PageBreakBefore>().Count() > 0)
                                    {
                                        Console.WriteLine("We have pagebreaks in the Paragraph;");
                                        foreach (DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore pbbf in para.Descendants<DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore>().ToArray())
                                        {
                                            para.RemoveChild<DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore>(pbbf);
                                        }
                                    }

                                    // remove any page breaks
                                    if (para.Elements<DocumentFormat.OpenXml.Wordprocessing.Break>().Count() > 0)
                                    {
                                        Console.WriteLine("We have breaks in the Paragraph;");
                                        foreach (DocumentFormat.OpenXml.Wordprocessing.Break brk in para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Break>().ToArray())
                                        {
                                            para.RemoveChild<DocumentFormat.OpenXml.Wordprocessing.Break>(brk);
                                        }
                                    }

                                    // correct any indentations
                                    if (pPr.Elements<Indentation>().Count() > 0)
                                    {
                                        Indentation ind = pPr.Elements<Indentation>().First();
                                        if (ind.Left != null)
                                        {
                                            if (Int32.Parse(ind.Left) > 0)
                                            {
                                                ind.Left = "720";
                                            }
                                        }
                                    }

                                }

                                // get a listof the text runs for the paragraph and adjust their settings as well
                                var runs = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>();
                                foreach (DocumentFormat.OpenXml.Wordprocessing.Run ru in runs.ToArray())
                                {
                                    // remove run styles
                                    if (ru.Elements<RunStyle>().Count() > 0)
                                    {
                                        ru.RemoveChild<RunStyle>(ru.Elements<RunStyle>().First());
                                    }

                                    // look for and create any run properties
                                    if (ru.Elements<RunProperties>().Count() == 0)
                                    {
                                        ru.PrependChild<RunProperties>(new RunProperties());
                                    }

                                    RunProperties ruProperty = ru.Elements<RunProperties>().First();
                                    ruProperty.RemoveAllChildren();

                                    RunFonts runFont = new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
                                    DocumentFormat.OpenXml.Wordprocessing.Color runColor = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "000000" };
                                    FontSize runSize = new FontSize { Val = new StringValue("22") };
                                    FontSizeComplexScript runSizeCS = new FontSizeComplexScript { Val = new StringValue("22") };

                                    ruProperty.AppendChild(runFont);
                                    ruProperty.AppendChild(runColor);
                                    ruProperty.AppendChild(runSize);
                                    ruProperty.AppendChild(runSizeCS);

                                    if (ru.Elements<PageBreakBefore>().Count() > 0)
                                    {
                                        Console.WriteLine("We have pagebreaks in the Run;");
                                        foreach (DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore pbbf in ru.Descendants<DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore>().ToArray())
                                        {
                                            ru.RemoveChild<DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore>(pbbf);
                                        }
                                    }

                                    if (ru.Elements<DocumentFormat.OpenXml.Wordprocessing.Break>().Count() > 0)
                                    {
                                        Console.WriteLine("We have breaks in the Run;");
                                        foreach (DocumentFormat.OpenXml.Wordprocessing.Break brk in ru.Descendants<DocumentFormat.OpenXml.Wordprocessing.Break>().ToArray())
                                        {
                                            ru.RemoveChild<DocumentFormat.OpenXml.Wordprocessing.Break>(brk);
                                        }
                                    }

                                }
                            }

                        }

                    }

                    var fileNameString = dir + "\\" + fn + "_processed.docx";
                    OpenXmlPackage docCloned = doc.Clone(fileNameString);

                    docCloned.Dispose();
                    doc.Dispose();

                    // remove the double spaces
                    using (WordprocessingDocument doc2 = WordprocessingDocument.Open(fileNameString, true))
                    {
                        SearchAndReplacer.SearchAndReplace(doc2, "     ", "\t", true);
                        SearchAndReplacer.SearchAndReplace(doc2, "  ", " ", true);
                    }
                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //AllocConsole();
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
    }
}