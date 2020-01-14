using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Speech.Recognition;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TextEditor {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        SpeechRecognitionEngine recog;
        public MainWindow() {
            InitializeComponent();

            // Get the speech culture (attempt to make recognition more accurate)
            var culture = (from r in SpeechRecognitionEngine.InstalledRecognizers()
                           where r.Culture.Equals(Thread.CurrentThread.CurrentCulture)
                           select r).FirstOrDefault();

            recog = new SpeechRecognitionEngine(culture);

            GrammarBuilder builder = new GrammarBuilder();
            builder.AppendDictation();
            Grammar grammar = new Grammar(builder);

            recog.LoadGrammar(grammar);
            recog.SetInputToDefaultAudioDevice();
            recog.SpeechRecognized += Recog_SpeechRecognized;

            // Get list of supported font families / initialize common font sizes
            var fontFamilies = Fonts.SystemFontFamilies.OrderBy(f => f.Source); // Order alphabetically
            familyBox.ItemsSource = fontFamilies;

            List<double> fontSizes = new List<double>() {8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 74, 96 };
            sizeBox.ItemsSource = fontSizes;


        }

        private void Recog_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
            string recognizedText = e.Result.Text;
            contentBox.Document.Blocks.Add(new Paragraph(new Run(recognizedText)));
        }

        private void exitMenuItem_Click(object sender, RoutedEventArgs e) {
            Application.Current.Shutdown();
        }

        private void familyBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (familyBox.SelectedItem != null) {
                contentBox.Selection.ApplyPropertyValue(Inline.FontFamilyProperty, familyBox.SelectedItem);
            }
        }

        private void sizeBox_TextChanged(object sender, TextChangedEventArgs e) {
            // Avoid crashing if the input to the size number for text is not a number
            try {
                double newSize = Convert.ToDouble(sizeBox.Text);
                contentBox.Selection.ApplyPropertyValue(Inline.FontSizeProperty, newSize);
            }
            catch (FormatException exc) {

            }
        }

        private void boldBtn_Click(object sender, RoutedEventArgs e) {
            bool isToggled = (sender as ToggleButton).IsChecked ?? false;

            if (isToggled)
                contentBox.Selection.ApplyPropertyValue(Inline.FontWeightProperty, FontWeights.Bold);
            else
                contentBox.Selection.ApplyPropertyValue(Inline.FontWeightProperty, FontWeights.Normal);
        }

        private void italBtn_Click(object sender, RoutedEventArgs e) {
            bool isToggled = (sender as ToggleButton).IsChecked ?? false;

            if (isToggled)
                contentBox.Selection.ApplyPropertyValue(Inline.FontStyleProperty, FontStyles.Italic);
            else
                contentBox.Selection.ApplyPropertyValue(Inline.FontStyleProperty, FontStyles.Normal);
        }

        private void underBtn_Click(object sender, RoutedEventArgs e) {
            bool isToggled = (sender as ToggleButton).IsChecked ?? false;

            if (isToggled) {
                contentBox.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, TextDecorations.Underline);
            }
            else {
                TextDecorationCollection textDecorations;
                (contentBox.Selection.GetPropertyValue(Inline.TextDecorationsProperty) as TextDecorationCollection)
                    .TryRemove(TextDecorations.Underline, out textDecorations);
                contentBox.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, textDecorations);
            }
        }

        private void contentBox_TextChanged(object sender, TextChangedEventArgs e) {
            // Keep track of # of characters
            int chars = (new TextRange(contentBox.Document.ContentStart, contentBox.Document.ContentEnd)).Text.Length;
            statusTextBlock.Text = $"Current Note Length: {chars} characters";
        }

        private void contentBox_SelectionChanged(object sender, RoutedEventArgs e) {
            var selectedWeight = contentBox.Selection.GetPropertyValue(Inline.FontWeightProperty);
            boldBtn.IsChecked = (selectedWeight != DependencyProperty.UnsetValue) && (selectedWeight.Equals(FontWeights.Bold));

            var selectedStyle = contentBox.Selection.GetPropertyValue(Inline.FontStyleProperty);
            italBtn.IsChecked = (selectedStyle != DependencyProperty.UnsetValue) && (selectedStyle.Equals(FontStyles.Italic));

            var selectedDecoration = contentBox.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            underBtn.IsChecked = (selectedDecoration != DependencyProperty.UnsetValue) && (selectedDecoration.Equals(TextDecorations.Underline));

            familyBox.SelectedItem = contentBox.Selection.GetPropertyValue(Inline.FontFamilyProperty);
            sizeBox.SelectedItem = contentBox.Selection.GetPropertyValue(Inline.FontSizeProperty);
        }

        private void speechBtn_Click(object sender, RoutedEventArgs e) {
            bool isToggled = (sender as ToggleButton).IsChecked ?? false; // If null -> make it false
            if (isToggled) {
                recog.RecognizeAsync(RecognizeMode.Multiple);
            }
            else {
                recog.RecognizeAsyncStop();
            }
        }

        private void newMenuItem_Click(object sender, RoutedEventArgs e) {
            contentBox.Document.Blocks.Clear();
        }

        private void openMenuItem_Click(object sender, RoutedEventArgs e) {
            OpenFileDialog file = new OpenFileDialog();
            if(file.ShowDialog() == true) {
                contentBox.Document.Blocks.Clear();
                contentBox.Document.Blocks.Add(new Paragraph(new Run(File.ReadAllText(file.FileName))));
            }
        }

        private void saveMenuItem_Click(object sender, RoutedEventArgs e) {
            SaveFileDialog file = new SaveFileDialog() {
                Filter = "Text Files(*.txt)|*.txt|All(*.*)|*"
            };
            if(file.ShowDialog() == true) {
                string savedText = new TextRange(contentBox.Document.ContentStart, contentBox.Document.ContentEnd).Text;
                File.WriteAllText(file.FileName, savedText);
            }
        }
    }
}
