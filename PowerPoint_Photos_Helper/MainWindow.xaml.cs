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
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Net;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

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

        private void rtbEditor_SelectionChanged(object sender, RoutedEventArgs e)
        {
            object temp = rtbEditor.Selection.GetPropertyValue(Inline.FontWeightProperty);
            btnBold.IsChecked = (temp != DependencyProperty.UnsetValue) && (temp.Equals(FontWeights.Bold));
            temp = rtbEditor.Selection.GetPropertyValue(Inline.FontFamilyProperty);
        }
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


        public string Folder { get; set; }
        public void SelectFolder(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            dlg.Title = "My Title";
            dlg.IsFolderPicker = true;

            dlg.AddToMostRecentlyUsedList = false;
            dlg.AllowNonFileSystemItems = false;
            dlg.EnsureFileExists = true;
            dlg.EnsurePathExists = true;
            dlg.EnsureReadOnly = false;
            dlg.EnsureValidNames = true;
            dlg.Multiselect = false;
            dlg.ShowPlacesList = true;

            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {
                Folder = dlg.FileName;
            }
        }

        private void SaveImage(string filename, ImageFormat format, string imageUrl)
        {
            WebClient client = new WebClient();
            Stream stream = client.OpenRead(imageUrl);
            Bitmap bitmap; bitmap = new Bitmap(stream);

            if (bitmap != null)
            {
                bitmap.Save(filename, format);
            }

            stream.Flush();
            stream.Close();
            client.Dispose();
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
            descriptionShape.TextBody.Text = rtbToString(rtbEditor);

            if (chk1.IsChecked == true && Folder != null && imageUrls.Count() > 0)
            {
                string fileName1 = Folder + "\\image1.png";
                SaveImage(fileName1, ImageFormat.Png, imageUrls[0]);
                Stream pictureStream = File.Open(fileName1, FileMode.Open);
                slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            }
            if (chk2.IsChecked == true && Folder != null && imageUrls.Count() > 1)
            {
                string fileName2 = Folder + "\\image2.png";
                SaveImage(fileName2, ImageFormat.Png, imageUrls[1]);
                Stream pictureStream = File.Open(fileName2, FileMode.Open);
                slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            }
            if (chk3.IsChecked == true && Folder != null && imageUrls.Count() > 2)
            {
                string fileName3 = Folder + "\\image3.png";
                SaveImage(fileName3, ImageFormat.Png, imageUrls[2]);
                Stream pictureStream = File.Open(fileName3, FileMode.Open);
                slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            }
            //Save the PowerPoint presentation
            powerpointDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            powerpointDoc.Close();

            //Open the PowerPoint presentation
            System.Diagnostics.Process.Start("Sample.pptx");
        }

        string rtbToString(RichTextBox rtb)

        {

            TextRange textRange = new TextRange(rtb.Document.ContentStart,

                rtb.Document.ContentEnd);

            return textRange.Text;

        }
        public void PhotosSearch(object sender, RoutedEventArgs e)
        {
            string query = Convert.ToString(TitleText.Text).Replace(" ", "+");
            string apiKey = "AIzaSyCXYc6Az6xQG7uwRfj-KgL0l93PgQu19JE";
            string url = "https://www.googleapis.com/customsearch/v1?key=" + apiKey + "&cx=b64b1dac81354d098&q=" + query;
            InitializeClient(url, imgDynamic1, imgDynamic2, imgDynamic3);
        }

        public static HttpClient ApiClient { get; set; }

        public static List<string> imageUrls = new List<string>();

        public async static void InitializeClient(string url, System.Windows.Controls.Image Img1, System.Windows.Controls.Image Img2, System.Windows.Controls.Image Img3)
        {
            ApiClient = new HttpClient();
            ApiClient.BaseAddress = new Uri(url);
            using (HttpResponseMessage response = await ApiClient.GetAsync(url))
            {
                string contentString = await response.Content.ReadAsStringAsync();
                dynamic parsedJson = JsonConvert.DeserializeObject(contentString);
                dynamic JsonItems = parsedJson.items;
                foreach (var item in JsonItems)
                {
                    if (item.pagemap.cse_image == null)
                    {

                    }
                    else
                    {
                        dynamic ImageUrl = item.pagemap.cse_image.ToString();
                        int first = ImageUrl.IndexOf(":")+3;
                        int last = ImageUrl.LastIndexOf("\"");
                        string rmv = ImageUrl.Remove(last).Remove(0, first);
                        if (System.IO.Path.HasExtension(rmv) == false)
                        {
                        }
                        else
                        {
                            imageUrls.Add(rmv);
                        }
                    }
                }
                int imageUrlsLength = imageUrls.Count();
                if (imageUrlsLength == 0)
                {
                }
                else if (imageUrlsLength == 1)
                {
                    Uri resourceUri1 = new Uri(imageUrls[0], UriKind.Absolute);
                    Img1.Source = new BitmapImage(resourceUri1);
                }
                else if (imageUrlsLength == 2)
                {
                    Uri resourceUri1 = new Uri(imageUrls[0], UriKind.Absolute);
                    Img1.Source = new BitmapImage(resourceUri1);
                    Uri resourceUri2 = new Uri(imageUrls[1], UriKind.Absolute);
                    Img2.Source = new BitmapImage(resourceUri2);
                }
                else if (imageUrlsLength >= 3)
                {
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
}