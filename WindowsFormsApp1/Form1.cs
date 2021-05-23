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

        private void buttonBold_Click_1(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = bodyText.SelectionFont;
            if (currentFont.Style != FontStyle.Bold)
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Bold);
            }
            else
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonUnderline_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = bodyText.SelectionFont;
            if (currentFont.Style != FontStyle.Underline)
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Underline);
            }
            else
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonItalics_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = bodyText.SelectionFont;
            if (currentFont.Style != FontStyle.Italic)
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Italic);
            }
            else
            {
                bodyText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void buttonBold1_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = titleText.SelectionFont;
            if (currentFont.Style != FontStyle.Bold)
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Bold);
            }
            else
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonUnderline1_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = titleText.SelectionFont;
            if (currentFont.Style != FontStyle.Underline)
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Underline);
            }
            else
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonItalics1_Click(object sender, EventArgs e)
        {
            System.Drawing.Font currentFont = titleText.SelectionFont;
            if (currentFont.Style != FontStyle.Italic)
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Italic);
            }
            else
            {
                titleText.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, FontStyle.Regular);
            }
        }

        private void buttonBGColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                bgColor.BackColor = colorDialog1.Color;
            }
        }

        private void buttonFontColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                fontColor.BackColor = colorDialog1.Color;
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            string titlePlainText = titleText.Text;
            string tempBodyText = "";
            string bodyPlainText = bodyText.Rtf;
            while (bodyPlainText.IndexOf("\\b") != -1)
            {
                int start = bodyPlainText.IndexOf("\\b") + 2;
                int end = bodyPlainText.IndexOf("\\b0");
                tempBodyText += bodyPlainText.Substring(start, end - start) + ' ';
                bodyPlainText = bodyPlainText.Substring(end + 3);
            }
            if(tempBodyText.Length != 0)
            {
                tempBodyText = tempBodyText.Substring(0, tempBodyText.Length - 1);
            }
            titlePlainText = titlePlainText + ' ' + tempBodyText;

            RunQueryAndDisplayResults(titlePlainText);
        }

        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            string plainTitleText = titleText.Text;
            string bodyPlainText = bodyText.Text;
            Color background = bgColor.BackColor;
            Color font = fontColor.BackColor;
            Application pptApplication = new Application();
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


    }
}
