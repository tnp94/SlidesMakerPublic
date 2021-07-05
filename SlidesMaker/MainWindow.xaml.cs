using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.Net;
using System.IO;
using System.Text.Json;

/* Not sure what the "3rd party" api restriction details are,
* but this one is from Microsoft and makes ppt slides.
*/
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace SlidesMaker
{
    public class BingImageResults
    {
        public IList<Dictionary<string, object>> value { get; set; }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>  
    public partial class MainWindow : Window
    {
        private PowerPoint.Application pptApp;
        private PowerPoint.Presentation presentation;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void getImageSuggestions(object sender, RoutedEventArgs e)
        {
            // Called upon clicking the "Get Image Suggestions" button
            List<string> keywords = getKeywords();
            wrapImagesFound.Children.Clear();
            populateImages(keywords);
        }

        private List<string> getKeywords()
        {
            /* Gets the keywords from the title and the bolded words in the text area
             * Returns a list of strings that are the keywords found
             * Args:
             *  none
             * Returns:
             *  List<string>: List of words in the title and bold words in the text area
             */

            List<string> keywords = new List<string>();
            // Add each title word as a keyword
            foreach(string word in txtTitle.Text.Split())
            {
                if (!string.IsNullOrWhiteSpace(word) && !keywords.Contains(word))
                    keywords.Add(word);
            }

            // Add each bolded word in the text area as a keyword
            foreach(Paragraph p in txtTextArea.Document.Blocks)
            {
                foreach(Inline inline in p.Inlines)
                {
                    if (inline.GetValue(FontWeightProperty).ToString() == "Bold")
                    {
                        string boldedSection = new TextRange(inline.ContentStart, inline.ContentEnd).Text;
                        List<string> boldWords = new List<string>(boldedSection.Split());
                        foreach (string keyword in boldWords)
                        {
                            if (!string.IsNullOrWhiteSpace(keyword) && !keywords.Contains(keyword))
                                keywords.Add(keyword);
                        }
                    }
                    else
                    {
                        string skippedWord = new TextRange(inline.ContentStart, inline.ContentEnd).Text;
                    }
                    
                }
            }
            return keywords;
        }

        private void populateImages(List<string> keywords)
        {
            /* Populates the wrap panel "wrapImagesFound" with images found from the internet
             * with the number of images being specified by the user in the slider or text box
             * with a minimum of 5 results and a maximum of 35 results.
             * (Bing only returns 35 max results in one query).
             * Args:
             *  keywords (List<string>): List of keywords to be searched in the Bing image api
             * Returns:
             *  none
             */
            if (keywords.Count < 1) // There were no keywords, display error and don't attempt the image search
            {
                Label errorMessage = new Label();
                errorMessage.Content = "Please enter a title or bold some words in the content area";
                wrapImagesFound.Children.Add(errorMessage);
                return;
            } 


            List<Image> imageList = new List<Image>();

            // Using documentation from https://docs.microsoft.com/en-us/azure/cognitive-services/bing-image-search/quickstarts/csharp
            string subscriptionKey = "{API Key Removed}";
            string uri = "https://api.bing.microsoft.com/v7.0/images/search";
            string searchTerm = string.Join(" ", keywords);
            var uriQuery = uri + "?q=" + Uri.EscapeDataString(searchTerm);

            WebRequest request = WebRequest.Create(uriQuery);
            request.Headers["Ocp-Apim-Subscription-Key"] = subscriptionKey;

            HttpWebResponse response = (HttpWebResponse)request.GetResponseAsync().Result;
            string jsonString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            BingImageResults json = JsonSerializer.Deserialize<BingImageResults>(jsonString);
            int maxResultsQuantity = json.value.Count < (int)sliderResultsCount.Value ? json.value.Count : (int)sliderResultsCount.Value;
            for(int i = 0; i < maxResultsQuantity; i++)
            {
                CheckBox checkBox = new CheckBox();
                Image image = new Image();
                string uriString = json.value[i]["contentUrl"].ToString();
                image.Source = new BitmapImage(new Uri(uriString));
                image.Width = 300;
                image.Height = 300;
                checkBox.Content = image;
                checkBox.Tag = uriString;
                wrapImagesFound.Children.Add(checkBox);
            }

        }
        private void createSlide(string title, string content, List<Image> images)
        {
            /* If there is no PowerPoint application open, opens PowerPoint. 
             * Then creates a slide with the selected images and inserts it at the beginning of the presentation.
             * Microsoft.Office.Interop.PowerPoint documentation (https://docs.microsoft.com/en-us/previous-versions/office/office-12/ff763170(v=office.12))
             * Args:
             *  title (string): The title of the slide to be created
             *  content (string): The content to be displayed in the body of the slide
             *  images (List<Image>): A list of images to be inserted onto the slide for the user
             * Returns:
             *  none
             */
            int titleYOffset = 0;
            int titleHeight = 100;

            int contentWidth = 400;
            int contentHeight = 600;

            int pictureWidth = 175;
            int picturesPerLine = 3;

            int contentXOffset = pictureWidth * picturesPerLine;
            int contentYOffset = titleHeight + titleYOffset;

            if (this.pptApp == null)
            {
                PowerPoint.Application newPptApp = new PowerPoint.Application();
                this.pptApp = newPptApp;
            }
            if (this.presentation == null)
            {
                PowerPoint.Presentation newPresentation = pptApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                this.presentation = newPresentation;
            }

            PowerPoint.CustomLayout layout;
            try
            {
                layout = presentation.SlideMaster.CustomLayouts[1];
                PowerPoint.Slides slides = presentation.Slides;
                PowerPoint.Slide slide = slides.AddSlide(1, layout);

                slide.Shapes.Title.Left = contentXOffset;
                slide.Shapes.Title.Width = contentWidth;
                slide.Shapes.Title.Top = titleYOffset;
                slide.Shapes.Title.Height = titleHeight;

                slide.Shapes.Title.TextFrame.TextRange.Text = title;

                PowerPoint.Shape contentBox = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, contentXOffset, contentYOffset, contentWidth, contentHeight);
                contentBox.TextFrame.TextRange.Text = content;

                for (int i = 0; i < images.Count; i++)
                {
                    PowerPoint.Shape picture = slide.Shapes.AddPicture(images[i].Source.ToString(), Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, (i % picturesPerLine) * pictureWidth, (i / picturesPerLine) * pictureWidth);
                    picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    picture.Width = pictureWidth;
                }

            } catch (System.Runtime.InteropServices.COMException e)
            {
                PowerPoint.Application newPptApp = new PowerPoint.Application();
                this.pptApp = newPptApp;
                PowerPoint.Presentation newPresentation = pptApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                this.presentation = newPresentation;
                MessageBoxResult errorMessage = MessageBox.Show(Application.Current.MainWindow, "PowerPoint application was closed, restarting. Please try again.");
            }
        }

        private void createSlideButton(object sender, RoutedEventArgs e)
        {
            /* Called upon clicking the "Add slide to presentation ->" button
             * Calls createSlide with the the information from the title, textarea, and checkbox fields
             */
            string content = new TextRange(txtTextArea.Document.ContentStart, txtTextArea.Document.ContentEnd).Text;
            List<Image> images = new List<Image>();
            foreach (CheckBox image in wrapImagesFound.Children)
            {
                if (image.IsChecked == true)
                {
                    images.Add((Image)image.Content);
                }
            }
            createSlide(txtTitle.Text, content, images);
        }
    }
}
