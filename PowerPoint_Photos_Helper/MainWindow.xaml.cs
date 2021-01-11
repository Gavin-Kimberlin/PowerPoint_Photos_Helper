using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Syncfusion.SfSkinManager;
using Syncfusion.Presentation;
using System.Net.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections;
using System.IO;

namespace PowerPoint_Photos_Helper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Fields
        private string currentVisualStyle;
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the current visual style.
        /// </summary>
        /// <value></value>
        /// <remarks></remarks>
        public string CurrentVisualStyle
        {
            get
            {
                return currentVisualStyle;
            }
            set
            {
                currentVisualStyle = value;
                OnVisualStyleChanged();
            }
        }
        #endregion
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += OnLoaded;
        }
        /// <summary>
        /// Called when [loaded].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            CurrentVisualStyle = "Metro";
        }
        /// <summary>
        /// On Visual Style Changed.
        /// </summary>
        /// <remarks></remarks>
        private void OnVisualStyleChanged()
        {
            VisualStyles visualStyle = VisualStyles.Default;
            Enum.TryParse(CurrentVisualStyle, out visualStyle);
            if (visualStyle != VisualStyles.Default)
            {
                SfSkinManager.ApplyStylesOnApplication = true;
                SfSkinManager.SetVisualStyle(this, visualStyle);
                SfSkinManager.ApplyStylesOnApplication = false;
            }
        }

        public void OnButtonClicked(object sender, RoutedEventArgs e)
        {
            //Create a new PowerPoint presentation
            IPresentation powerpointDoc = Presentation.Create();

            //Add a blank slide to the presentation
            ISlide slide = powerpointDoc.Slides.Add(SlideLayoutType.Blank);

            //Add a textbox to the slide
            IShape shape = slide.AddTextBox(400, 100, 500, 100);

            //Add a text to the textbox
            string title = Convert.ToString(TitleText.Text);
            shape.TextBody.AddParagraph(title);

            //Add description content to the slide by adding a new TextBox
            IShape descriptionShape = slide.AddTextBox(53.22, 141.73, 874.19, 77.70);
            descriptionShape.TextBody.Text = BodyText.Text;

            //Save the PowerPoint presentation
            powerpointDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            powerpointDoc.Close();

            //Open the PowerPoint presentation
            System.Diagnostics.Process.Start("Sample.pptx");

        }

        public void PhotosSearch(object sender, RoutedEventArgs e)
        {
            string query = Convert.ToString(TitleText.Text).Replace(" ", "+");
            string apiKey = "AIzaSyAEquGPHOTnWnmOTbPKkzq7mLgUrUlBdyw";
            string url = "https://www.googleapis.com/customsearch/v1?key=" + apiKey + "&cx=b64b1dac81354d098&q=" + query;
            InitializeClient(url, imgDynamic1, imgDynamic2, imgDynamic3);
        }

        public static HttpClient ApiClient { get; set; }

        //private static Image imgDynamic1;



        public static List<string> imageUrls = new List<string>();
        public async static void InitializeClient(string url, Image Img1, Image Img2, Image Img3)
        {
            ApiClient = new HttpClient();
            ApiClient.BaseAddress = new Uri(url);
            using (HttpResponseMessage response = await ApiClient.GetAsync(url))
            {
                string contentString = await response.Content.ReadAsStringAsync();
                dynamic parsedJson = JsonConvert.DeserializeObject(contentString);
                foreach (var item in parsedJson.items)
                {
                    if (item.pagemap.cse_image == null)
                    {

                    }
                    else
                    {
                        dynamic ImageUrl = item.pagemap.cse_image.ToString();
                        int first = ImageUrl.IndexOf("h");
                        int last = ImageUrl.LastIndexOf("g") + 1;
                        string rmv = ImageUrl.Remove(last).Remove(0, first);
                        imageUrls.Add(rmv);
                    }
                }
                Uri resourceUri1 = new Uri(imageUrls[0], UriKind.Absolute);
                Img1.Source = new BitmapImage(resourceUri1);

                Uri resourceUri2 = new Uri(imageUrls[1], UriKind.Absolute);
                Img2.Source = new BitmapImage(resourceUri2);

                Uri resourceUri3 = new Uri(imageUrls[2], UriKind.Absolute);
                Img3.Source = new BitmapImage(resourceUri3);
            }
        }
    }
}