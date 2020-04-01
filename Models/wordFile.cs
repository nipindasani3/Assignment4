using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using CSV.Models;
using System.IO;

namespace Office_File_Formats.Models
{
    class wordFile
    {
        private static Drawing AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                       "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to the body. The element should be in 
            // a DocumentFormat.OpenXml.Wordprocessing.Run.
            //wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
            return element;
        }

        public static void CreateWordprocessingDocument(string filepath)
        {
            List<Student> studentlist = Methods._readCSV(Constants.Locations.StudentCSVPath);
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {

                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                string[] getFileList = Directory.GetFiles(Constants.Locations.StudentImageFolder);
                foreach (var student in studentlist)
                {
                    run.AppendChild(new Text($"Hello, My Name is '{ student.FirstName } { student.LastName}'"));
                    run.AppendChild(new Break());
                    run.AppendChild(new Text($"My StudentID is '{ student.StudentId}'"));
                    run.AppendChild(new Break());
                    run.AppendChild(new Text($"My Birthdate is '{ student.DateOfBirth}'"));
                    run.AppendChild(new Break());
                    run.AppendChild(new Break());
                    foreach (String files in getFileList)
                    {
                        String str = Path.GetFileName(files);
                        if (str.StartsWith(student.StudentId))
                        {
                            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                            using (FileStream stream = new FileStream(files, FileMode.Open))
                            {
                                imagePart.FeedData(stream);
                            }
                            run.AppendChild(AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart)));
                        }
                    }
                    run.Append(new Break() { Type = BreakValues.Page });
                }
            }

        }
    }
}
