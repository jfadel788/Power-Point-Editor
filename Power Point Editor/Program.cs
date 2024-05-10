//By Jawad Fadel
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//Please Install the pacakge Syncfusion.Presentation.Net.core
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1
//Open an existing PowerPoint presentation
using Syncfusion.Drawing;
using Syncfusion.Presentation;
using System.Data.SqlTypes;
using System.Runtime.CompilerServices;
using System.Xml;

IPresentation pptxDoc = Presentation.Open(new FileStream(path: "C:\\Users\\USER\\source\\repos\\Power Point Editor\\Power Point Editor\\pp.pptx", FileMode.Open));
//Gets the first slide from the PowerPoint presentation
ISlide slide = pptxDoc.Slides[0];
var count=slide.Shapes.Count;

var shapes= slide.Shapes;
//Gets the first shape of the slide
IFont font;
slide.Shapes[0].Left = 300;
IShape ss = slide.Shapes[0] as IShape;
ss.TextBody.Text = "Output Slide";
 
for(int j=0;j<5;j++)
{
    IShape ishape = (IShape)slide.Shapes[5];
    if(ishape.TextBody.Text=="Step 1" || ishape.TextBody.Text == "Step 2"|| ishape.TextBody.Text == "Begin"|| ishape.TextBody.Text == "Done")
    {
        slide.Shapes.Remove(ishape);
    }
}


IShape s1 = slide.Shapes[1] as IShape;
s1.TextBody.Text = "Begin";
IShape s2 = slide.Shapes[2] as IShape;
s2.TextBody.Text = "Step 1";
IShape s3 = slide.Shapes[3] as IShape;
s3.TextBody.Text = "Step 2";
IShape s4 = slide.Shapes[4] as IShape;
s4.TextBody.Text = "Done";
for (int i = 1; i <= 4; i++)
{
    
    IShape c = slide.Shapes[i] as IShape;

    c.TextBody.Paragraphs[0].TextParts[0].Font.Color =ColorObject.White;
    slide.Shapes[i].Top = slide.Shapes[0].Top + 100;
    slide.Shapes[i].Width = 220;
    slide.Shapes[i].Height = 105;
    if (i > 1)
    {
        slide.Shapes[i].Left = (slide.Shapes[i - 1].Left + 180);

    }
}

for (int i = 0; i < 4; i++)
{

    IShape shape = slide.Shapes[i + 5] as IShape;
    var iSlideItem = slide.Shapes[i].GetType();
    for (int j = 0; j < 4; j++)
    {
        IParagraph paragraph = shape.TextBody.Paragraphs[j];
        //Retrieves the first TextPart of the shape.

        ITextPart textPart = paragraph.TextParts[0];
        textPart.Font.Underline = TextUnderlineType.None;
        textPart.Font.Italic = false;
        textPart.Font.Bold = false;
        paragraph.ListFormat.Type = ListType.Bulleted;
        paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
        paragraph.ListFormat.Size = 70;
        paragraph.ListFormat.FontName = "Symbol";


    }




    Console.WriteLine(iSlideItem.Name);
}

//Save the PowerPoint presentation as stream

using (var stream = new FileStream("C:\\Users\\USER\\source\\repos\\Power Point Editor\\Power Point Editor\\output1.pptx", FileMode.Create))
{
    pptxDoc.Save(stream);
    stream.Position = 0;
}
//outputStream.Flush();
//outputStream.Dispose();
////Close the PowerPoint presentation
pptxDoc.Close();