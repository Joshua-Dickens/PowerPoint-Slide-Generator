using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Web;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using System.IO;
using System.Drawing.Imaging;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void bodyText_TextChanged(object sender, EventArgs e)
        {

        }

        // Gives the user the Ability to bold words in the rich text box for the body
        private void buttonBold_Click_1(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = bodyText.SelectionFont;
            if (currentFont.Style != FontStyle.Bold)
            {
                bodyText.SelectionFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, FontStyle.Bold);
            }
            else
            {
                bodyText.SelectionFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
        // Closes the Application
        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Runs the Image search
        private void buttonSearch_Click(object sender, EventArgs e)
        {
            string titlePlainText = titleText.Text;
            string tempBodyText = "";
            string bodyPlainText = bodyText.Rtf;
            // Extracts the bolded keywords of the body
            while (bodyPlainText.IndexOf("\\b") != -1)
            {
                int start = bodyPlainText.IndexOf("\\b") + 2;
                int end = bodyPlainText.IndexOf("\\b0");
                tempBodyText += bodyPlainText.Substring(start, end - start) + ' ';
                bodyPlainText = bodyPlainText.Substring(end + 3);
            }
            // Checks if there were any bold words at all
            if(tempBodyText.Length != 0)
            {
                tempBodyText = tempBodyText.Substring(0, tempBodyText.Length - 1);
            }
            titlePlainText = titlePlainText + ' ' + tempBodyText;
            //Runs the code to search and display content
            RunQueryAndDisplayResults(titlePlainText);
        }

        // Runs the slide generation
        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            string plainTitleText = titleText.Text;
            string bodyPlainText = bodyText.Text;
            bool[] checkBoxes = {checkBox1.Checked, checkBox2.Checked, checkBox3.Checked, checkBox4.Checked, checkBox5.Checked, checkBox6.Checked, checkBox7.Checked, checkBox8.Checked, checkBox9.Checked, checkBox10.Checked };

            string filepath = "";
            //Prompts the user to select the powerpoint
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 0;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filepath = openFileDialog.FileName;
                }
            }
            progressBar1.Value = 0;
            // Creates the new slide with title text and body text
            InsertNewSlide(filepath, 2, plainTitleText, bodyPlainText);
            progressBar1.Value = 30;
            // Adds the images to the slide
            int used = 0;
            var imageCopy = pictureBox1.Image;
            string imageFilePath = filepath.Substring(0, filepath.LastIndexOf(@"\") + 1) + "tempImage.png";
            progressBar1.Value = 50;
            for (int i = 0; i < 10; i++)
            {
                progressBar1.Value += 5;
                if(used < 5 && checkBoxes[i])
                {
                    
                    switch (i + 1)
                    {
                        case 1:
                            imageCopy = pictureBox1.Image;
                            break;
                        case 2:
                            imageCopy = pictureBox2.Image;
                            break;
                        case 3:
                            imageCopy = pictureBox3.Image;
                            break;
                        case 4:
                            imageCopy = pictureBox4.Image;
                            break;
                        case 5:
                            imageCopy = pictureBox5.Image;
                            break;
                        case 6:
                            imageCopy = pictureBox6.Image;
                            break;
                        case 7:
                            imageCopy = pictureBox7.Image;
                            break;
                        case 8:
                            imageCopy = pictureBox8.Image;
                            break;
                        case 9:
                            imageCopy = pictureBox9.Image;
                            break;
                        case 10:
                            imageCopy = pictureBox10.Image;
                            break;
                    }
                    imageCopy.Save(@imageFilePath, ImageFormat.Png);
                    AddImage(filepath, @imageFilePath, 500000 + 2250000 * used, 4500000);
                    used++;
                }
            }
            // Deletes the image storage
            File.Delete(imageFilePath);
        }



        private void RunQueryAndDisplayResults(string userQuery)
        {
            try
            {
                // Create a query
                var client = new HttpClient();
                client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "dc8a13278a8a4018a4cd19b0c8f32344");
                var queryString = HttpUtility.ParseQueryString(string.Empty);
                queryString["q"] = userQuery;
                queryString["responseFilter"] = "images";
                var query = "https://api.bing.microsoft.com/v7.0/search?" + queryString;

                // Run the query
                HttpResponseMessage httpResponseMessage = client.GetAsync(query).Result;

                // Deserialize the response content
                var responseContentString = httpResponseMessage.Content.ReadAsStringAsync().Result;
                Newtonsoft.Json.Linq.JObject responseObjects = Newtonsoft.Json.Linq.JObject.Parse(responseContentString);

                // Handle success and error codes
                if (httpResponseMessage.IsSuccessStatusCode)
                {
                    DisplayAllRankedResults(responseObjects);
                }
                else
                {
                    Console.WriteLine($"HTTP error status code: {httpResponseMessage.StatusCode.ToString()}");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void DisplayAllRankedResults(Newtonsoft.Json.Linq.JObject responseObjects)
        {
            string[] rankingGroups = new string[] { "pole", "mainline", "sidebar" };

            // Loop through the ranking groups in priority order
            foreach (string rankingName in rankingGroups)
            {
                Newtonsoft.Json.Linq.JToken rankingResponseItems = responseObjects.SelectToken($"rankingResponse.{rankingName}.items");
                if (rankingResponseItems != null)
                {
                    foreach (Newtonsoft.Json.Linq.JObject rankingResponseItem in rankingResponseItems)
                    {
                        Newtonsoft.Json.Linq.JToken resultIndex;
                        rankingResponseItem.TryGetValue("resultIndex", out resultIndex);
                        var answerType = rankingResponseItem.Value<string>("answerType");
                        DisplaySpecificResults(resultIndex, responseObjects.SelectToken("images.value"), "Image", "thumbnailUrl");
                        
                    }
                }
            }
        }

        private void DisplaySpecificResults(Newtonsoft.Json.Linq.JToken resultIndex, Newtonsoft.Json.Linq.JToken items, string title, params string[] fields)
        {
            int x = 0;
            if (resultIndex == null)
            {
                foreach (Newtonsoft.Json.Linq.JToken item in items)
                {
                    x++;
                    displayPicture(item, fields, x);
                }
            }
        }

        private void displayPicture(Newtonsoft.Json.Linq.JToken item, string[] fields, int increment)
        {
            // Assaigns the first 10 results to their respective picture boxes
            switch (increment)
            {
                case 1:
                    pictureBox1.ImageLocation = item[fields[0]].ToString();
                    break;
                case 2:
                    pictureBox2.ImageLocation = item[fields[0]].ToString();
                    break;
                case 3:
                    pictureBox3.ImageLocation = item[fields[0]].ToString();
                    break;
                case 4:
                    pictureBox4.ImageLocation = item[fields[0]].ToString();
                    break;
                case 5:
                    pictureBox5.ImageLocation = item[fields[0]].ToString();
                    break;
                case 6:
                    pictureBox6.ImageLocation = item[fields[0]].ToString();
                    break;
                case 7:
                    pictureBox7.ImageLocation = item[fields[0]].ToString();
                    break;
                case 8:
                    pictureBox8.ImageLocation = item[fields[0]].ToString();
                    break;
                case 9:
                    pictureBox9.ImageLocation = item[fields[0]].ToString();
                    break;
                case 10:
                    pictureBox10.ImageLocation = item[fields[0]].ToString();
                    break;
            }
            
        }

        public static void InsertNewSlide(string presentationFile, int position, string slideTitle, string body)
        {
            // Open the source document as read/write. 
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                // Pass the source document and the position and title of the slide to be inserted to the next method.
                InsertNewSlide(presentationDocument, position, slideTitle, body);
            }
        }

        // Insert the specified slide into the presentation at the specified position.
        public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle, string body)
        {

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (slideTitle == null)
            {
                throw new ArgumentNullException("slideTitle");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            // Declare and instantiate the title shape of the new slide.
            Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));

            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
            drawingObjectId++;

            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the body shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = body })));

            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

        public static void AddImage(string file, string image, int x, int y)
        {
            using (var presentation = PresentationDocument.Open(file, true))
            {
                var slidePart = presentation
                    .PresentationPart
                    .SlideParts
                    .Last();

                var part = slidePart
                    .AddImagePart(ImagePartType.Png);

                using (var stream = File.OpenRead(image))
                {
                    part.FeedData(stream);
                }


                var tree = slidePart
                    .Slide
                    .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
                    .First();

                var picture = new DocumentFormat.OpenXml.Presentation.Picture();

                picture.NonVisualPictureProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties();
                picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
                {
                    Name = "My Shape",
                    Id = (UInt32)tree.ChildElements.Count - 1
                });

                var nonVisualPictureDrawingProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties();
                nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
                {
                    NoChangeAspect = true
                });
                picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
                picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

                var blipFill = new DocumentFormat.OpenXml.Presentation.BlipFill();
                var blip1 = new DocumentFormat.OpenXml.Drawing.Blip()
                {
                    Embed = slidePart.GetIdOfPart(part)
                };
                var blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
                var blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension()
                {
                    Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                };
                var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
                {
                    Val = false
                };
                useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                blipExtension1.Append(useLocalDpi1);
                blipExtensionList1.Append(blipExtension1);
                blip1.Append(blipExtensionList1);
                var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
                stretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
                blipFill.Append(blip1);
                blipFill.Append(stretch);
                picture.Append(blipFill);

                picture.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
                picture.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
                picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
                {
                    X = x,
                    Y = y,
                });
                picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
                {
                    Cx = 2000000,
                    Cy = 2000000,
                });
                picture.ShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry
                {
                    Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                });

                tree.Append(picture);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        
    }
}
