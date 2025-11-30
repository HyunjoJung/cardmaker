// Copyright (c) 2025 Business Card Maker Contributors
// Licensed under the Apache License, Version 2.0

using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace BusinessCardMaker.Core.Services.Template;

/// <summary>
/// Service for creating sample PowerPoint business card templates
/// </summary>
public class TemplateService : ITemplateService
{
    private const long EMU_PER_INCH = 914400;

    public byte[] CreateBasicTemplate()
    {
        using var stream = new MemoryStream();
        using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            CreatePresentationParts(document);

            var slidePart = CreateSlide(document.PresentationPart!);

            // Add text boxes with placeholders
            AddTextBox(slidePart, "{name}", 0.5, 0.5, 3.0, 0.5, 24, true);
            AddTextBox(slidePart, "{position}", 0.5, 1.1, 3.0, 0.4, 14, false);
            AddTextBox(slidePart, "{department}", 0.5, 1.6, 3.0, 0.4, 12, false);
            AddTextBox(slidePart, "Email: {email}", 0.5, 2.3, 4.0, 0.3, 11, false);
            AddTextBox(slidePart, "Mobile: {mobile}", 0.5, 2.7, 4.0, 0.3, 11, false);
            AddTextBox(slidePart, "Phone: {phone}", 0.5, 3.1, 4.0, 0.3, 11, false);

            // Company info
            AddTextBox(slidePart, "{company}", 5.5, 0.5, 3.5, 0.5, 16, true);
            AddTextBox(slidePart, "Your Company Address", 5.5, 1.1, 3.5, 0.8, 10, false);

            slidePart.Slide.Save();
        }

        return stream.ToArray();
    }

    public byte[] CreateQRCodeTemplate()
    {
        using var stream = new MemoryStream();
        using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            CreatePresentationParts(document);

            var slidePart = CreateSlide(document.PresentationPart!);

            // Add text boxes with placeholders
            AddTextBox(slidePart, "{name}", 0.5, 0.5, 3.0, 0.5, 24, true);
            AddTextBox(slidePart, "{position}", 0.5, 1.1, 3.0, 0.4, 14, false);
            AddTextBox(slidePart, "{department}", 0.5, 1.6, 3.0, 0.4, 12, false);
            AddTextBox(slidePart, "Email: {email}", 0.5, 2.3, 3.5, 0.3, 11, false);
            AddTextBox(slidePart, "Mobile: {mobile}", 0.5, 2.7, 3.5, 0.3, 11, false);

            // QR code placeholder
            AddTextBox(slidePart, "{qrcode}", 7.0, 2.0, 1.5, 1.5, 10, false);

            // Company info
            AddTextBox(slidePart, "{company}", 5.0, 0.5, 3.5, 0.5, 16, true);
            AddTextBox(slidePart, "Your Company Address", 5.0, 1.1, 3.5, 0.8, 10, false);

            slidePart.Slide.Save();
        }

        return stream.ToArray();
    }

    private void CreatePresentationParts(PresentationDocument document)
    {
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        // Create slide master and layout
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

        slideMasterPart.SlideMaster = new SlideMaster(
            new P.CommonSlideData(new P.ShapeTree()),
            new P.ColorMap()
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new P.SlideLayoutIdList());

        slideLayoutPart.SlideLayout = new SlideLayout(
            new P.CommonSlideData(new P.ShapeTree()),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        var slideLayoutIdList = slideMasterPart.SlideMaster.SlideLayoutIdList!;
        var slideLayoutId = new P.SlideLayoutId { Id = 2147483649, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) };
        slideLayoutIdList.Append(slideLayoutId);

        presentationPart.Presentation.SlideIdList = new SlideIdList();
        presentationPart.Presentation.SlideSize = new SlideSize { Cx = 9144000, Cy = 6858000 }; // 10" x 7.5"

        presentationPart.Presentation.Save();
    }

    private SlidePart CreateSlide(PresentationPart presentationPart)
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide(
            new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        var slideIdList = presentationPart.Presentation.SlideIdList!;
        uint maxSlideId = 256;
        var slideId = new SlideId { Id = maxSlideId, RelationshipId = presentationPart.GetIdOfPart(slidePart) };
        slideIdList.Append(slideId);

        return slidePart;
    }

    private void AddTextBox(SlidePart slidePart, string text, double x, double y, double width, double height, int fontSize, bool bold)
    {
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
        uint shapeId = (uint)(shapeTree.ChildElements.Count + 1);

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = (long)(x * EMU_PER_INCH), Y = (long)(y * EMU_PER_INCH) },
                    new A.Extents { Cx = (long)(width * EMU_PER_INCH), Cy = (long)(height * EMU_PER_INCH) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(
                        new A.RunProperties
                        {
                            Language = "en-US",
                            FontSize = fontSize * 100,
                            Bold = bold
                        },
                        new A.Text(text)))));

        shapeTree.Append(shape);
    }
}
