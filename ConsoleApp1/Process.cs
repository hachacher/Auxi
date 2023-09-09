using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using System;
using System.Collections.Generic;

namespace ConsoleApp1
{
    public class Process
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private PresentationDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = PresentationDocument.Open(filePath, true))
            {
                ChangeParts();
            }
        }

        private void ChangeParts()
        { 
            BuildUriPartDictionary();
           
            ChangeSlidePart1(((SlidePart)UriPartDictionary["/ppt/slides/slide1.xml"]));
            
        } 
        private void BuildUriPartDictionary()
        {
            System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
            queue.Enqueue(document);
            while (queue.Count > 0)
            {
                foreach (var part in queue.Dequeue().Parts)
                {
                    if (!UriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString()))
                    {
                        UriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                        queue.Enqueue(part.OpenXmlPart);
                    }
                }
            }
        }

        private void ChangeSlidePart1(SlidePart slidePart1)
        {
            Slide slide1 = slidePart1.Slide;

            CommonSlideData commonSlideData1 = slide1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();
            CommonSlideDataExtensionList commonSlideDataExtensionList1 = commonSlideData1.GetFirstChild<CommonSlideDataExtensionList>();

            Shape shape1 = shapeTree1.GetFirstChild<Shape>();
            Shape shape2 = shapeTree1.Elements<Shape>().ElementAt(1);
            Shape shape3 = shapeTree1.Elements<Shape>().ElementAt(2);
            Shape shape4 = shapeTree1.Elements<Shape>().ElementAt(3);
            Shape shape5 = shapeTree1.Elements<Shape>().ElementAt(4);
            Shape shape6 = shapeTree1.Elements<Shape>().ElementAt(5);
            Shape shape7 = shapeTree1.Elements<Shape>().ElementAt(6);
            Shape shape8 = shapeTree1.Elements<Shape>().ElementAt(7);
            Shape shape9 = shapeTree1.Elements<Shape>().ElementAt(8);
            Shape shape10 = shapeTree1.Elements<Shape>().ElementAt(9);
            Shape shape11 = shapeTree1.Elements<Shape>().ElementAt(10);
            Shape shape12 = shapeTree1.Elements<Shape>().ElementAt(11);
            Shape shape13 = shapeTree1.Elements<Shape>().ElementAt(12);
            List<Shape> shapelist = shapeTree1.Elements<Shape>().ToList();
            TextBody textBody1 = shape1.GetFirstChild<TextBody>();

            A.Paragraph paragraph1 = textBody1.GetFirstChild<A.Paragraph>();

            A.Run run1 = paragraph1.GetFirstChild<A.Run>();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            paragraph1.InsertBefore(paragraphProperties1, run1);

            A.RunProperties runProperties1 = run1.GetFirstChild<A.RunProperties>();
            A.Text text1 = run1.GetFirstChild<A.Text>();

            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Beirut", PitchFamily = 2, CharacterSet = -78 };
            runProperties1.Append(latinFont1);

            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "Beirut", PitchFamily = 2, CharacterSet = -78 };
            runProperties1.Append(complexScriptFont1);
            text1.Text = "Output Slide";


            ShapeProperties shapeProperties1 = shape2.GetFirstChild<ShapeProperties>();
            TextBody textBody2 = shape2.GetFirstChild<TextBody>();

            A.Transform2D transform2D1 = shapeProperties1.GetFirstChild<A.Transform2D>();

            A.Offset offset1 = transform2D1.GetFirstChild<A.Offset>();
            A.Extents extents1 = transform2D1.GetFirstChild<A.Extents>();
            offset1.X = 458568L;
            extents1.Cx = 2775204L;

            A.Paragraph paragraph2 = textBody2.GetFirstChild<A.Paragraph>();

            A.EndParagraphRunProperties endParagraphRunProperties1 = paragraph2.GetFirstChild<A.EndParagraphRunProperties>();

            A.Run run2 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text2 = new A.Text();
            shapelist.Remove(shape2);
            text2.Text = CheckforTextBox(shapelist,offset1.X,offset1.Y);

            run2.Append(runProperties2);
            run2.Append(text2);
            paragraph2.InsertBefore(run2, endParagraphRunProperties1);

            endParagraphRunProperties1.Remove();

            ShapeProperties shapeProperties2 = shape3.GetFirstChild<ShapeProperties>();
            TextBody textBody3 = shape3.GetFirstChild<TextBody>();

            A.Transform2D transform2D2 = shapeProperties2.GetFirstChild<A.Transform2D>();

            A.Offset offset2 = transform2D2.GetFirstChild<A.Offset>();
            A.Extents extents2 = transform2D2.GetFirstChild<A.Extents>();
            offset2.X = 2616031L;
            offset2.Y = 1712339L;
            extents2.Cx = 2775204L;
            extents2.Cy = 1446935L;

            A.Paragraph paragraph3 = textBody3.GetFirstChild<A.Paragraph>();

            A.EndParagraphRunProperties endParagraphRunProperties2 = paragraph3.GetFirstChild<A.EndParagraphRunProperties>();

            A.Run run3 = new A.Run();
            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text3 = new A.Text();
            shapelist.Remove(shape3);
            text3.Text = CheckforTextBox(shapelist, offset2.X, offset2.Y);

            run3.Append(runProperties3);
            run3.Append(text3);
            paragraph3.InsertBefore(run3, endParagraphRunProperties2);

            endParagraphRunProperties2.Remove();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill1.Append(schemeColor1);

            endParagraphRunProperties3.Append(solidFill1);

            paragraph4.Append(paragraphProperties2);
            paragraph4.Append(endParagraphRunProperties3);
            textBody3.Append(paragraph4);

            NonVisualShapeProperties nonVisualShapeProperties1 = shape4.GetFirstChild<NonVisualShapeProperties>();
            ShapeProperties shapeProperties3 = shape4.GetFirstChild<ShapeProperties>();
            ShapeStyle shapeStyle1 = shape4.GetFirstChild<ShapeStyle>();
            TextBody textBody4 = shape4.GetFirstChild<TextBody>();

            NonVisualDrawingProperties nonVisualDrawingProperties1 = nonVisualShapeProperties1.GetFirstChild<NonVisualDrawingProperties>();
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = nonVisualShapeProperties1.GetFirstChild<NonVisualShapeDrawingProperties>();
            nonVisualDrawingProperties1.Id = (UInt32Value)7U;
            nonVisualDrawingProperties1.Name = "TextBox 6";

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = nonVisualDrawingProperties1.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = nonVisualDrawingPropertiesExtensionList1.GetFirstChild<A.NonVisualDrawingPropertiesExtension>();

            OpenXmlUnknownElement openXmlUnknownElement1 = nonVisualDrawingPropertiesExtension1.GetFirstChild<OpenXmlUnknownElement>();

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{16092594-6C62-E547-8EB0-D18626DD5308}\" />");
            nonVisualDrawingPropertiesExtension1.InsertBefore(openXmlUnknownElement2, openXmlUnknownElement1);

            openXmlUnknownElement1.Remove();
            nonVisualShapeDrawingProperties1.TextBox = true;

            A.Transform2D transform2D3 = shapeProperties3.GetFirstChild<A.Transform2D>();
            A.PresetGeometry presetGeometry1 = shapeProperties3.GetFirstChild<A.PresetGeometry>();

            A.Offset offset3 = transform2D3.GetFirstChild<A.Offset>();
            A.Extents extents3 = transform2D3.GetFirstChild<A.Extents>();
            offset3.X = 408013L;
            offset3.Y = 3700095L;
            extents3.Cx = 1527662L;
            extents3.Cy = 1200329L;
            presetGeometry1.Preset = A.ShapeTypeValues.Rectangle;

            A.NoFill noFill1 = new A.NoFill();
            shapeProperties3.Append(noFill1);

            shapeStyle1.Remove();

            A.BodyProperties bodyProperties1 = textBody4.GetFirstChild<A.BodyProperties>();
            A.Paragraph paragraph5 = textBody4.GetFirstChild<A.Paragraph>();
            bodyProperties1.Anchor = null;
            bodyProperties1.Wrap = A.TextWrappingValues.None;

            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();
            bodyProperties1.Append(shapeAutoFit1);

            A.ParagraphProperties paragraphProperties3 = paragraph5.GetFirstChild<A.ParagraphProperties>();
            A.EndParagraphRunProperties endParagraphRunProperties4 = paragraph5.GetFirstChild<A.EndParagraphRunProperties>();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont1 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet1 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties4.Append(bulletFont1);
            paragraphProperties4.Append(characterBullet1);
            paragraph5.InsertBefore(paragraphProperties4, paragraphProperties3);

            A.Run run4 = new A.Run();
            A.RunProperties runProperties4 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text4 = new A.Text();
            text4.Text = "Start here";

            run4.Append(runProperties4);
            run4.Append(text4);
            paragraph5.InsertBefore(run4, paragraphProperties3);

            paragraphProperties3.Remove();
            endParagraphRunProperties4.Remove();

            A.Paragraph paragraph6 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont2 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet2 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties5.Append(bulletFont2);
            paragraphProperties5.Append(characterBullet2);

            A.Run run5 = new A.Run();
            A.RunProperties runProperties5 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text5 = new A.Text();
            text5.Text = "Maybe not";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph6.Append(paragraphProperties5);
            paragraph6.Append(run5);
            textBody4.Append(paragraph6);

            A.Paragraph paragraph7 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont3 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet3 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties6.Append(bulletFont3);
            paragraphProperties6.Append(characterBullet3);

            A.Run run6 = new A.Run();
            A.RunProperties runProperties6 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text6 = new A.Text();
            text6.Text = "Lobster roll";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph7.Append(paragraphProperties6);
            paragraph7.Append(run6);
            textBody4.Append(paragraph7);

            A.Paragraph paragraph8 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont4 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet4 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties7.Append(bulletFont4);
            paragraphProperties7.Append(characterBullet4);

            A.Run run7 = new A.Run();
            A.RunProperties runProperties7 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text7 = new A.Text();
            text7.Text = "C#";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph8.Append(paragraphProperties7);
            paragraph8.Append(run7);
            textBody4.Append(paragraph8);

            NonVisualShapeProperties nonVisualShapeProperties2 = shape5.GetFirstChild<NonVisualShapeProperties>();
            ShapeProperties shapeProperties4 = shape5.GetFirstChild<ShapeProperties>();
            TextBody textBody5 = shape5.GetFirstChild<TextBody>();

            NonVisualDrawingProperties nonVisualDrawingProperties2 = nonVisualShapeProperties2.GetFirstChild<NonVisualDrawingProperties>();
            nonVisualDrawingProperties2.Id = (UInt32Value)18U;
            nonVisualDrawingProperties2.Name = "Arrow: Chevron 8";

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = nonVisualDrawingProperties2.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = nonVisualDrawingPropertiesExtensionList2.GetFirstChild<A.NonVisualDrawingPropertiesExtension>();

            OpenXmlUnknownElement openXmlUnknownElement3 = nonVisualDrawingPropertiesExtension2.GetFirstChild<OpenXmlUnknownElement>();

            OpenXmlUnknownElement openXmlUnknownElement4 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{6F5DBADA-2C1B-6944-B0E2-C4CBB3BDD33D}\" />");
            nonVisualDrawingPropertiesExtension2.InsertBefore(openXmlUnknownElement4, openXmlUnknownElement3);

            openXmlUnknownElement3.Remove();

            A.Transform2D transform2D4 = shapeProperties4.GetFirstChild<A.Transform2D>();

            A.Offset offset4 = transform2D4.GetFirstChild<A.Offset>();
            A.Extents extents4 = transform2D4.GetFirstChild<A.Extents>();
            offset4.X = 4794237L;
            offset4.Y = 1716258L;
            extents4.Cx = 2775204L;
            extents4.Cy = 1446935L;

            A.Paragraph paragraph9 = textBody5.GetFirstChild<A.Paragraph>();

            A.EndParagraphRunProperties endParagraphRunProperties5 = paragraph9.GetFirstChild<A.EndParagraphRunProperties>();

            A.Run run8 = new A.Run();
            A.RunProperties runProperties8 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text8 = new A.Text();
            shapelist.Remove(shape5);
            text8.Text =CheckforTextBox(shapelist,offset4.X,offset4.Y);

            run8.Append(runProperties8);
            run8.Append(text8);
            paragraph9.InsertBefore(run8, endParagraphRunProperties5);

            endParagraphRunProperties5.Remove();

            A.Paragraph paragraph10 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill2.Append(schemeColor2);

            endParagraphRunProperties6.Append(solidFill2);

            paragraph10.Append(paragraphProperties8);
            paragraph10.Append(endParagraphRunProperties6);
            textBody5.Append(paragraph10);

            NonVisualShapeProperties nonVisualShapeProperties3 = shape6.GetFirstChild<NonVisualShapeProperties>();
            ShapeProperties shapeProperties5 = shape6.GetFirstChild<ShapeProperties>();
            TextBody textBody6 = shape6.GetFirstChild<TextBody>();

            NonVisualDrawingProperties nonVisualDrawingProperties3 = nonVisualShapeProperties3.GetFirstChild<NonVisualDrawingProperties>();
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = nonVisualShapeProperties3.GetFirstChild<NonVisualShapeDrawingProperties>();
            nonVisualDrawingProperties3.Id = (UInt32Value)19U;
            nonVisualDrawingProperties3.Name = "Arrow: Chevron 8";

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList3 = nonVisualDrawingProperties3.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension3 = nonVisualDrawingPropertiesExtensionList3.GetFirstChild<A.NonVisualDrawingPropertiesExtension>();

            OpenXmlUnknownElement openXmlUnknownElement5 = nonVisualDrawingPropertiesExtension3.GetFirstChild<OpenXmlUnknownElement>();

            OpenXmlUnknownElement openXmlUnknownElement6 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{6F99AC8F-10CB-5F4C-B53B-0E206FBDC708}\" />");
            nonVisualDrawingPropertiesExtension3.InsertBefore(openXmlUnknownElement6, openXmlUnknownElement5);

            openXmlUnknownElement5.Remove();
            nonVisualShapeDrawingProperties2.TextBox = null;

            A.Transform2D transform2D5 = shapeProperties5.GetFirstChild<A.Transform2D>();
            A.PresetGeometry presetGeometry2 = shapeProperties5.GetFirstChild<A.PresetGeometry>();
            A.NoFill noFill2 = shapeProperties5.GetFirstChild<A.NoFill>();

            A.Offset offset5 = transform2D5.GetFirstChild<A.Offset>();
            A.Extents extents5 = transform2D5.GetFirstChild<A.Extents>();
            offset5.X = 6972443L;
            offset5.Y = 1716258L;
            extents5.Cx = 2775204L;
            extents5.Cy = 1446935L;
            presetGeometry2.Preset = A.ShapeTypeValues.Chevron;

            noFill2.Remove();

            ShapeStyle shapeStyle2 = new ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade1 = new A.Shade() { Val = 50000 };

            schemeColor3.Append(shade1);

            lineReference1.Append(schemeColor3);

            A.FillReference fillReference1 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor4);

            A.EffectReference effectReference1 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor5);

            A.FontReference fontReference1 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference1.Append(schemeColor6);

            shapeStyle2.Append(lineReference1);
            shapeStyle2.Append(fillReference1);
            shapeStyle2.Append(effectReference1);
            shapeStyle2.Append(fontReference1);
            shape6.InsertBefore(shapeStyle2, textBody6);

            A.BodyProperties bodyProperties2 = textBody6.GetFirstChild<A.BodyProperties>();
            A.Paragraph paragraph11 = textBody6.GetFirstChild<A.Paragraph>();
            bodyProperties2.Wrap = null;
            bodyProperties2.Anchor = A.TextAnchoringTypeValues.Center;

            A.ShapeAutoFit shapeAutoFit2 = bodyProperties2.GetFirstChild<A.ShapeAutoFit>();

            shapeAutoFit2.Remove();

            A.Run run9 = paragraph11.GetFirstChild<A.Run>();

            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            paragraph11.InsertBefore(paragraphProperties9, run9);

            A.RunProperties runProperties9 = run9.GetFirstChild<A.RunProperties>();
            A.Text text9 = run9.GetFirstChild<A.Text>();
            runProperties9.FontSize = null;
            shapelist.Remove(shape6);

            text9.Text = CheckforTextBox(shapelist,offset5.X,offset5.Y);


            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill3.Append(schemeColor7);

            endParagraphRunProperties7.Append(solidFill3);
            paragraph11.Append(endParagraphRunProperties7);

            NonVisualShapeProperties nonVisualShapeProperties4 = shape7.GetFirstChild<NonVisualShapeProperties>();
            ShapeProperties shapeProperties6 = shape7.GetFirstChild<ShapeProperties>();
            TextBody textBody7 = shape7.GetFirstChild<TextBody>();

            NonVisualDrawingProperties nonVisualDrawingProperties4 = nonVisualShapeProperties4.GetFirstChild<NonVisualDrawingProperties>();
            nonVisualDrawingProperties4.Id = (UInt32Value)20U;
            nonVisualDrawingProperties4.Name = "TextBox 19";

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList4 = nonVisualDrawingProperties4.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension4 = nonVisualDrawingPropertiesExtensionList4.GetFirstChild<A.NonVisualDrawingPropertiesExtension>();

            OpenXmlUnknownElement openXmlUnknownElement7 = nonVisualDrawingPropertiesExtension4.GetFirstChild<OpenXmlUnknownElement>();

            OpenXmlUnknownElement openXmlUnknownElement8 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{BFEE9ABA-C571-7E41-BA45-6FF07623AB24}\" />");
            nonVisualDrawingPropertiesExtension4.InsertBefore(openXmlUnknownElement8, openXmlUnknownElement7);

            openXmlUnknownElement7.Remove();

            A.Transform2D transform2D6 = shapeProperties6.GetFirstChild<A.Transform2D>();

            A.Offset offset6 = transform2D6.GetFirstChild<A.Offset>();
            A.Extents extents6 = transform2D6.GetFirstChild<A.Extents>();
            offset6.X = 2756961L;
            offset6.Y = 3694807L;
            extents6.Cx = 1527662L;
            extents6.Cy = 1200329L;

            A.BodyProperties bodyProperties3 = textBody7.GetFirstChild<A.BodyProperties>();
            A.Paragraph paragraph12 = textBody7.GetFirstChild<A.Paragraph>();
            bodyProperties3.Wrap = A.TextWrappingValues.None;

            A.Run run10 = paragraph12.GetFirstChild<A.Run>();

            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont5 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet5 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties10.Append(bulletFont5);
            paragraphProperties10.Append(characterBullet5);
            paragraph12.InsertBefore(paragraphProperties10, run10);

            A.RunProperties runProperties10 = run10.GetFirstChild<A.RunProperties>();
            A.Text text10 = run10.GetFirstChild<A.Text>();
            runProperties10.FontSize = null;
            text10.Text = "Start here";


            A.Paragraph paragraph13 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont6 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet6 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties11.Append(bulletFont6);
            paragraphProperties11.Append(characterBullet6);

            A.Run run11 = new A.Run();
            A.RunProperties runProperties11 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text11 = new A.Text();
            text11.Text = "Maybe not";

            run11.Append(runProperties11);
            run11.Append(text11);

            paragraph13.Append(paragraphProperties11);
            paragraph13.Append(run11);
            textBody7.Append(paragraph13);

            A.Paragraph paragraph14 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont7 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet7 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties12.Append(bulletFont7);
            paragraphProperties12.Append(characterBullet7);

            A.Run run12 = new A.Run();
            A.RunProperties runProperties12 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text12 = new A.Text();
            text12.Text = "Lobster roll";

            run12.Append(runProperties12);
            run12.Append(text12);

            paragraph14.Append(paragraphProperties12);
            paragraph14.Append(run12);
            textBody7.Append(paragraph14);

            A.Paragraph paragraph15 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties13 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont8 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet8 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties13.Append(bulletFont8);
            paragraphProperties13.Append(characterBullet8);

            A.Run run13 = new A.Run();
            A.RunProperties runProperties13 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text13 = new A.Text();
            text13.Text = "C#";

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph15.Append(paragraphProperties13);
            paragraph15.Append(run13);
            textBody7.Append(paragraph15);

            NonVisualShapeProperties nonVisualShapeProperties5 = shape8.GetFirstChild<NonVisualShapeProperties>();
            ShapeProperties shapeProperties7 = shape8.GetFirstChild<ShapeProperties>();
            TextBody textBody8 = shape8.GetFirstChild<TextBody>();

            NonVisualDrawingProperties nonVisualDrawingProperties5 = nonVisualShapeProperties5.GetFirstChild<NonVisualDrawingProperties>();
            nonVisualDrawingProperties5.Id = (UInt32Value)21U;
            nonVisualDrawingProperties5.Name = "TextBox 20";

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList5 = nonVisualDrawingProperties5.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension5 = nonVisualDrawingPropertiesExtensionList5.GetFirstChild<A.NonVisualDrawingPropertiesExtension>();

            OpenXmlUnknownElement openXmlUnknownElement9 = nonVisualDrawingPropertiesExtension5.GetFirstChild<OpenXmlUnknownElement>();

            OpenXmlUnknownElement openXmlUnknownElement10 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C7DC06DC-AEB9-9C46-8D33-CA310FFA0813}\" />");
            nonVisualDrawingPropertiesExtension5.InsertBefore(openXmlUnknownElement10, openXmlUnknownElement9);

            openXmlUnknownElement9.Remove();

            A.Transform2D transform2D7 = shapeProperties7.GetFirstChild<A.Transform2D>();

            A.Offset offset7 = transform2D7.GetFirstChild<A.Offset>();
            A.Extents extents7 = transform2D7.GetFirstChild<A.Extents>();
            offset7.X = 5105909L;
            offset7.Y = 3694806L;
            extents7.Cx = 1527662L;
            extents7.Cy = 1200329L;

            A.BodyProperties bodyProperties4 = textBody8.GetFirstChild<A.BodyProperties>();
            A.Paragraph paragraph16 = textBody8.GetFirstChild<A.Paragraph>();
            bodyProperties4.Wrap = A.TextWrappingValues.None;

            A.Run run14 = paragraph16.GetFirstChild<A.Run>();

            A.ParagraphProperties paragraphProperties14 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont9 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet9 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties14.Append(bulletFont9);
            paragraphProperties14.Append(characterBullet9);
            paragraph16.InsertBefore(paragraphProperties14, run14);

            A.RunProperties runProperties14 = run14.GetFirstChild<A.RunProperties>();
            A.Text text14 = run14.GetFirstChild<A.Text>();
            runProperties14.FontSize = null;
            text14.Text = "Start here";


            A.Paragraph paragraph17 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties15 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont10 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet10 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties15.Append(bulletFont10);
            paragraphProperties15.Append(characterBullet10);

            A.Run run15 = new A.Run();
            A.RunProperties runProperties15 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text15 = new A.Text();
            text15.Text = "Maybe not";

            run15.Append(runProperties15);
            run15.Append(text15);

            paragraph17.Append(paragraphProperties15);
            paragraph17.Append(run15);
            textBody8.Append(paragraph17);

            A.Paragraph paragraph18 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties16 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont11 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet11 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties16.Append(bulletFont11);
            paragraphProperties16.Append(characterBullet11);

            A.Run run16 = new A.Run();
            A.RunProperties runProperties16 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text16 = new A.Text();
            text16.Text = "Lobster roll";

            run16.Append(runProperties16);
            run16.Append(text16);

            paragraph18.Append(paragraphProperties16);
            paragraph18.Append(run16);
            textBody8.Append(paragraph18);

            A.Paragraph paragraph19 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties17 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont12 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet12 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties17.Append(bulletFont12);
            paragraphProperties17.Append(characterBullet12);

            A.Run run17 = new A.Run();
            A.RunProperties runProperties17 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text17 = new A.Text();
            text17.Text = "C#";

            run17.Append(runProperties17);
            run17.Append(text17);

            paragraph19.Append(paragraphProperties17);
            paragraph19.Append(run17);
            textBody8.Append(paragraph19);

            NonVisualShapeProperties nonVisualShapeProperties6 = shape9.GetFirstChild<NonVisualShapeProperties>();
            ShapeProperties shapeProperties8 = shape9.GetFirstChild<ShapeProperties>();
            TextBody textBody9 = shape9.GetFirstChild<TextBody>();

            NonVisualDrawingProperties nonVisualDrawingProperties6 = nonVisualShapeProperties6.GetFirstChild<NonVisualDrawingProperties>();
            nonVisualDrawingProperties6.Id = (UInt32Value)22U;
            nonVisualDrawingProperties6.Name = "TextBox 21";

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList6 = nonVisualDrawingProperties6.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension6 = nonVisualDrawingPropertiesExtensionList6.GetFirstChild<A.NonVisualDrawingPropertiesExtension>();

            OpenXmlUnknownElement openXmlUnknownElement11 = nonVisualDrawingPropertiesExtension6.GetFirstChild<OpenXmlUnknownElement>();

            OpenXmlUnknownElement openXmlUnknownElement12 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{515CF982-AFD7-BA4A-ACD6-41A01DBB1386}\" />");
            nonVisualDrawingPropertiesExtension6.InsertBefore(openXmlUnknownElement12, openXmlUnknownElement11);

            openXmlUnknownElement11.Remove();

            A.Transform2D transform2D8 = shapeProperties8.GetFirstChild<A.Transform2D>();

            A.Offset offset8 = transform2D8.GetFirstChild<A.Offset>();
            A.Extents extents8 = transform2D8.GetFirstChild<A.Extents>();
            offset8.X = 7143548L;
            offset8.Y = 3694805L;
            extents8.Cx = 1527662L;
            extents8.Cy = 1200329L;

            A.BodyProperties bodyProperties5 = textBody9.GetFirstChild<A.BodyProperties>();
            A.Paragraph paragraph20 = textBody9.GetFirstChild<A.Paragraph>();
            bodyProperties5.Wrap = A.TextWrappingValues.None;

            A.Run run18 = paragraph20.GetFirstChild<A.Run>();

            A.ParagraphProperties paragraphProperties18 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont13 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet13 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties18.Append(bulletFont13);
            paragraphProperties18.Append(characterBullet13);
            paragraph20.InsertBefore(paragraphProperties18, run18);

            A.RunProperties runProperties18 = run18.GetFirstChild<A.RunProperties>();
            A.Text text18 = run18.GetFirstChild<A.Text>();
            runProperties18.FontSize = null;
            text18.Text = "Start here";


            A.Paragraph paragraph21 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties19 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont14 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet14 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties19.Append(bulletFont14);
            paragraphProperties19.Append(characterBullet14);

            A.Run run19 = new A.Run();
            A.RunProperties runProperties19 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text19 = new A.Text();
            text19.Text = "Maybe not";

            run19.Append(runProperties19);
            run19.Append(text19);

            paragraph21.Append(paragraphProperties19);
            paragraph21.Append(run19);
            textBody9.Append(paragraph21);

            A.Paragraph paragraph22 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties20 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont15 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet15 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties20.Append(bulletFont15);
            paragraphProperties20.Append(characterBullet15);

            A.Run run20 = new A.Run();
            A.RunProperties runProperties20 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text20 = new A.Text();
            text20.Text = "Lobster roll";

            run20.Append(runProperties20);
            run20.Append(text20);

            paragraph22.Append(paragraphProperties20);
            paragraph22.Append(run20);
            textBody9.Append(paragraph22);

            A.Paragraph paragraph23 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties21 = new A.ParagraphProperties() { LeftMargin = 285750, Indent = -285750 };
            A.BulletFont bulletFont16 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet16 = new A.CharacterBullet() { Char = "•" };

            paragraphProperties21.Append(bulletFont16);
            paragraphProperties21.Append(characterBullet16);

            A.Run run21 = new A.Run();
            A.RunProperties runProperties21 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text21 = new A.Text();
            text21.Text = "C#";

            run21.Append(runProperties21);
            run21.Append(text21);

            paragraph23.Append(paragraphProperties21);
            paragraph23.Append(run21);
            textBody9.Append(paragraph23);

            shape10.Remove();
            shape11.Remove();
            shape12.Remove();
            shape13.Remove();

            CommonSlideDataExtension commonSlideDataExtension1 = commonSlideDataExtensionList1.GetFirstChild<CommonSlideDataExtension>();

            P14.CreationId creationId1 = commonSlideDataExtension1.GetFirstChild<P14.CreationId>();
            creationId1.Val = (UInt32Value)3777547941U;
        }



        public string CheckforTextBox(List<Shape> ShapeList, Int64Value xT, Int64Value yT)
        {
            foreach(Shape s in ShapeList)
            {
                ShapeProperties shapeProperties1 = s.GetFirstChild<ShapeProperties>();

                 
                A.Transform2D transform2D1 = shapeProperties1.GetFirstChild<A.Transform2D>();
                if (transform2D1 != null)
                {
                    
                    if (check(transform2D1.Offset.X, transform2D1.Offset.Y, transform2D1.Extents.Cx, transform2D1.Extents.Cy, xT, yT))
                    {
                        return s.InnerText;
                    }

                }
            }
            
            return "";
        }
        static bool check(Int64Value x, Int64Value y, Int64Value Cx,
                      Int64Value Cy , Int64Value xT, Int64Value yT)
        {

            if ((xT.Value.CompareTo(x.Value) > 0) && (xT.Value.CompareTo(x.Value) < 0) && (yT.Value.CompareTo(y.Value) > 0) && (yT.Value.CompareTo(y.Value) < 0))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        
    }
}
