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
using Microsoft.Office.Interop.PowerPoint;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Ink;
using Point = System.Windows.Point;
using System.Threading;
using System.Windows.Interop;
using System.Windows.Threading;

namespace PowerpointJabber
{
    public partial class SimplePenWindow : Window
    {
        public List<EditingButton> pens;
        private List<SlideIndicator> slides = new List<SlideIndicator>();
        private EditingButton currentPen;
        private Dictionary<int, bool> clickAdvanceStates = new Dictionary<int, bool>();
        DispatcherTimer backgroundPolling;
        public IntPtr HWND;
        public SimplePenWindow()
        {
            InitializeComponent();
            pens = new List<EditingButton>
                {
                    new EditingButton(EditingButton.EditingType.Pen,"black",System.Windows.Media.Brushes.Black),
                    new EditingButton(EditingButton.EditingType.Pen,"blue",System.Windows.Media.Brushes.Blue),
                    new EditingButton(EditingButton.EditingType.Pen,"red",System.Windows.Media.Brushes.Red),
                    new EditingButton(EditingButton.EditingType.Pen,"green",System.Windows.Media.Brushes.Green),
                    new EditingButton(EditingButton.EditingType.Pen,"yellow",System.Windows.Media.Brushes.Yellow),
                    new EditingButton(EditingButton.EditingType.Pen,"orange",System.Windows.Media.Brushes.Orange),
                    new EditingButton(EditingButton.EditingType.Pen,"white",System.Windows.Media.Brushes.White),
                    new EditingButton(EditingButton.EditingType.Eraser,"eraser",System.Windows.Media.Brushes.Transparent)
                };
            populateSlidesAdvanceDictionary();
            currentPen = pens[0];
            PensControl.Items.Clear();
            PensControl.ItemsSource = pens;
            foreach (var slide in slides)
                clickAdvanceStates.Add(slide.slideId, slide.clickAdvance);
            if (shouldWorkaroundClickAdvance)
                setClickAdvanceOnAllSlides(false);
            if (backgroundPolling == null)
                backgroundPolling = backgroundPollingTimer();
        }
        private DispatcherTimer backgroundPollingTimer()
        {
            if (this.Dispatcher == null) return null;
            return new DispatcherTimer(TimeSpan.FromMilliseconds(250), DispatcherPriority.Background, delegate
            {
                try
                {
                    if (this == null || ThisAddIn.instance.Application == null || ThisAddIn.instance.Application.SlideShowWindows.Count < 1)
                        return;
                    var state = WindowsInteropFunctions.getAppropriateViewData();
                    if (HWND == null || !((int)HWND > 0))
                    {
                        HwndSource source = (HwndSource)HwndSource.FromVisual(this);
                        HWND = source.Handle;
                    }
                    if ((this.WindowState != WindowState.Minimized) != state.isVisible)
                        this.WindowState = state.isVisible ? WindowState.Normal : WindowState.Minimized;
                    if (!Double.IsNaN(state.X) && this.Left != state.X)
                        this.Left = state.X;
                    if (WindowsInteropFunctions.presenterActive)
                    {
                        if (!Double.IsNaN(state.Y) && !Double.IsNaN(state.Height))
                        {
                            var newY = (state.Y + (state.Height * 0.06));
                            if (this.Top != newY)
                                this.Top = newY;
                        }
                    }
                    else this.Top = 0;
                    if (!Double.IsNaN(state.Height) && state.Height > 0 && ViewboxContainer.ActualHeight != state.Height)
                        ViewboxContainer.Height = state.Height * 0.6;
                    if (ThisAddIn.instance != null
                        && ThisAddIn.instance.Application != null
                        && ThisAddIn.instance.Application.ActivePresentation != null
                        && ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow != null
                        && ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View != null
                        && ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.PointerColor != null)
                    {
                        switch (ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.PointerType)
                        {
                            case PpSlideShowPointerType.ppSlideShowPointerPen:
                                int currentColour = ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.PointerColor.RGB;
                                bool penHasBeenFound = false;
                                foreach (var pen in pens)
                                {
                                    if (pen.type == EditingButton.EditingType.Pen && pen.RGBAasInt == currentColour)
                                    {
                                        currentPen = pen;
                                        selectPen(pen);
                                        penHasBeenFound = true;
                                    }
                                    if (!penHasBeenFound)
                                        selectPen(null);
                                }
                                break;
                            case PpSlideShowPointerType.ppSlideShowPointerEraser:
                                foreach (var pen in pens)
                                {
                                    if (pen.type == EditingButton.EditingType.Eraser)
                                        selectPen(pen);
                                }
                                break;
                            default:
                                foreach (var pen in pens)
                                    selectPen(null);
                                break;
                        }
                    }
                }
                catch (Exception) { }
            }, this.Dispatcher);
        }

        private bool shouldWorkaroundClickAdvance { get { return WindowsInteropFunctions.presenterActive && !pptVersionIs2010; } }
        private bool pptVersionIs2010
        {
            get
            {
                return Double.Parse(ThisAddIn.instance.Application.Version) > 12;
            }
        }
        private void populateSlidesAdvanceDictionary()
        {
            slides.Clear();
            foreach (Slide slide in ThisAddIn.instance.Application.ActivePresentation.Slides)
            {
                slides.Add(new SlideIndicator(slide.SlideID));
            }
        }
        private void setClickAdvanceOnAllSlides(bool state)
        {
            if (shouldWorkaroundClickAdvance)
                foreach (var slide in slides)
                    slide.setClickAdvance(state);
        }
        private void ReFocusPresenter()
        {
            WindowsInteropFunctions.BringAppropriateViewToFront();
        }
        private void selectPen(EditingButton button)
        {
            if (pens == null || pens.Count == 0) return;
            if (button == null)
                foreach (var pen in pens)
                    pen.Selected = false;
            foreach (var pen in pens)
            {
                if (pen == button)
                {
                    pen.Selected = true;
                }
                else pen.Selected = false;
            }
        }
        private void Pen(object sender, RoutedEventArgs e)
        {
            var internalCurrentPen = pens.Where(c => c.name == ((FrameworkElement)sender).Tag.ToString()).FirstOrDefault();
            Logger.Info("{0} pen selected by SimplePens", internalCurrentPen.name);
            switch (internalCurrentPen.type)
            {
                case EditingButton.EditingType.Eraser:
                    ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.PointerType = PpSlideShowPointerType.ppSlideShowPointerEraser;
                    break;
                case EditingButton.EditingType.Pen:
                    ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.PointerColor.RGB = internalCurrentPen.RGBAasInt;
                    ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.PointerType = PpSlideShowPointerType.ppSlideShowPointerPen;
                    break;
                case EditingButton.EditingType.Selector:
                    ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.PointerType = PpSlideShowPointerType.ppSlideShowPointerArrow;
                    break;
                default:
                    ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.PointerType = PpSlideShowPointerType.ppSlideShowPointerNone;
                    break;
            }
            ReFocusPresenter();
        }
        private void EndSlideShow(object sender, RoutedEventArgs e)
        {
            if (backgroundPolling != null)
            {
                try
                {
                    backgroundPolling.Stop();
                    backgroundPolling = null;
                }
                catch (Exception) { }
            }
            if (ThisAddIn.instance != null && ThisAddIn.instance.Application != null && ThisAddIn.instance.Application.SlideShowWindows != null && ThisAddIn.instance.Application.SlideShowWindows.Count > 0)
                try
                {
                    ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.Exit();
                }
                catch (Exception) { }
            if (this != null)
            {
                try
                {
                    this.Close();
                }
                catch (Exception) { }
            }
        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                if (backgroundPolling != null)
                {
                    backgroundPolling.Stop();
                    backgroundPolling = null;
                }
                foreach (var slide in slides)
                {
                    slide.clickAdvance = clickAdvanceStates[slide.slideId];
                }
                ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.Exit();
            }
            catch (Exception)
            {
            }
        }
        private void closeApplication(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                if (backgroundPolling != null)
                {
                    backgroundPolling.Stop();
                    backgroundPolling = null;
                }
                Close();
            }
            catch (Exception) { }
        }

        private void AddPage(object sender, RoutedEventArgs e)
        {
            Logger.Info("Page added by SimplePens");
            CustomLayout newSlide;
            if (ThisAddIn.instance.Application.ActivePresentation.SlideMaster.CustomLayouts.Count > 0)
                newSlide = ThisAddIn.instance.Application.ActivePresentation.SlideMaster.CustomLayouts[1];
            else
                newSlide = ThisAddIn.instance.Application.ActivePresentation.Slides[0].CustomLayout;
            if (newSlide != null)
            {
                var newSlideIndex = ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition + 1;
                ThisAddIn.instance.Application.ActivePresentation.Slides.AddSlide(newSlideIndex, newSlide);
                ThisAddIn.instance.Application.ActivePresentation.Slides[newSlideIndex].Layout = PpSlideLayout.ppLayoutBlank;
                ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.Activate();
                ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.GotoSlide(newSlideIndex);
                var slideIndicator = new SlideIndicator(ThisAddIn.instance.Application.ActivePresentation.Slides[newSlideIndex].SlideID);
                slides.Add(slideIndicator);
                slideIndicator.clickAdvance = false;
                clickAdvanceStates.Add(slideIndicator.slideId, false);
                ReFocusPresenter();
            }
        }
        public class EditingButton : DependencyObject
        {
            public EditingButton(EditingType Type, string Name, System.Windows.Media.SolidColorBrush Color)
            {
                penColour = Color;
                name = Name;
                type = Type;
                generateRGBAsInt();
                generateDrawnPenPreview();
                generateBrushPreviewPoints();
            }
            public bool Selected
            {
                get { return (bool)GetValue(SelectedProperty); }
                set { SetValue(SelectedProperty, value); }
            }
            public static readonly DependencyProperty SelectedProperty =
                DependencyProperty.Register("Selected", typeof(bool), typeof(EditingButton), new UIPropertyMetadata(false));

            public System.Windows.Media.Brush HighlightColour { get; private set; }
            public enum EditingType { Pen, Eraser, Selector }
            public string name { get; private set; }
            public EditingType type { get; private set; }
            public System.Windows.Media.SolidColorBrush penColour { get; private set; }
            private int R { get { return penColour.Color.R; } }
            private int G { get { return penColour.Color.G; } }
            private int B { get { return penColour.Color.B; } }
            private int A { get { return penColour.Color.A; } }
            private int cachedRGBAsInt;
            private bool RGBAsIntHasBeenCached = false;
            private void generateRGBAsInt()
            {
                cachedRGBAsInt = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(A, R, G, B));
                RGBAsIntHasBeenCached = true;
            }
            public int RGBAasInt
            {
                get
                {
                    if (!RGBAsIntHasBeenCached)
                    {
                        generateRGBAsInt();
                    }
                    return cachedRGBAsInt;
                }
            }
            public string tooltip
            {
                get
                {
                    string result = "";
                    switch (type)
                    {
                        case EditingType.Pen:
                            result = "Draw with a " + name + " pen.";
                            break;
                        case EditingType.Eraser:
                            result = "Erase";
                            break;
                        case EditingType.Selector:
                            result = "Arrow";
                            break;
                        default:
                            result = "";
                            break;
                    }
                    return result;
                }
            }
            private StrokeCollection cachedDrawnPenPreviewStroke;
            public StrokeCollection DrawnPenPreviewStroke
            {
                get
                {
                    if (cachedDrawnPenPreviewStroke == null)
                        generateDrawnPenPreview();
                    return cachedDrawnPenPreviewStroke;
                }
            }

            private void generateDrawnPenPreview()
            {
                cachedDrawnPenPreviewStroke = new StrokeCollection(
                        new[]{
                            new Stroke(
                                new StylusPointCollection(
                                    new StylusPoint[]{
                                        new StylusPoint(30.6666666666667,90,0.5f),
                                        new StylusPoint(32.6666666666667,91.3333333333333,0.5f),
                                        new StylusPoint(33.6666666666667,91.6666666666667,0.5f),
                                        new StylusPoint(35,92,0.5f),
                                        new StylusPoint(35.6666666666667,92.3333333333333,0.5f),
                                        new StylusPoint(36.3333333333333,92.6666666666667,0.5f),
                                        new StylusPoint(37.3333333333333,93,0.5f),
                                        new StylusPoint(38,93.3333333333333,0.5f),
                                        new StylusPoint(39,93.6666666666667,0.5f),
                                        new StylusPoint(40.3333333333333,94,0.5f),
                                        new StylusPoint(41.3333333333333,94.3333333333333,0.5f),
                                        new StylusPoint(42.6666666666667,94.3333333333333,0.5f),
                                        new StylusPoint(43.6666666666667,94.6666666666667,0.5f),
                                        new StylusPoint(45.3333333333333,95,0.5f),
                                        new StylusPoint(46.6666666666667,95.3333333333333,0.5f),
                                        new StylusPoint(48,95.3333333333333,0.5f),
                                        new StylusPoint(49.3333333333333,95.3333333333333,0.5f),
                                        new StylusPoint(51,95.6666666666667,0.5f),
                                        new StylusPoint(52.6666666666667,95.6666666666667,0.5f),
                                        new StylusPoint(54,95.3333333333333,0.5f),
                                        new StylusPoint(55.6666666666667,95.3333333333333,0.5f),
                                        new StylusPoint(57.3333333333333,95,0.5f),
                                        new StylusPoint(59,94.6666666666667,0.5f),
                                        new StylusPoint(60.6666666666667,94.3333333333333,0.5f),
                                        new StylusPoint(62.3333333333333,94,0.5f),
                                        new StylusPoint(64.3333333333333,93.3333333333333,0.5f),
                                        new StylusPoint(65.6666666666667,93,0.5f),
                                        new StylusPoint(67.3333333333333,92.3333333333333,0.5f),
                                        new StylusPoint(69,91.6666666666667,0.5f),
                                        new StylusPoint(70.6666666666667,91,0.5f),
                                        new StylusPoint(72,90.3333333333333,0.5f),
                                        new StylusPoint(73.6666666666667,89.3333333333333,0.5f),
                                        new StylusPoint(75,88.6666666666667,0.5f),
                                        new StylusPoint(76.3333333333333,87.6666666666667,0.5f),
                                        new StylusPoint(77.3333333333333,86.6666666666667,0.5f),
                                        new StylusPoint(78.6666666666667,85.6666666666667,0.5f),
                                        new StylusPoint(79.6666666666667,84.6666666666667,0.5f),
                                        new StylusPoint(80.6666666666667,83.6666666666667,0.5f),
                                        new StylusPoint(81.6666666666667,82.3333333333333,0.5f),
                                        new StylusPoint(82.6666666666667,81,0.5f),
                                        new StylusPoint(83.3333333333333,80,0.5f),
                                        new StylusPoint(84,78.6666666666667,0.5f),
                                        new StylusPoint(84.3333333333333,77.3333333333333,0.5f),
                                        new StylusPoint(85,76,0.5f),
                                        new StylusPoint(85.3333333333333,74.6666666666667,0.5f),
                                        new StylusPoint(85.6666666666667,73,0.5f),
                                        new StylusPoint(86,71.6666666666667,0.5f),
                                        new StylusPoint(86,70.3333333333333,0.5f),
                                        new StylusPoint(86,69,0.5f),
                                        new StylusPoint(86,67.6666666666667,0.5f),
                                        new StylusPoint(85.6666666666667,66.3333333333333,0.5f),
                                        new StylusPoint(85.6666666666667,65,0.5f),
                                        new StylusPoint(85.3333333333333,63.6666666666667,0.5f),
                                        new StylusPoint(85,62.3333333333333,0.5f),
                                        new StylusPoint(84.3333333333333,61,0.5f),
                                        new StylusPoint(83.6666666666667,59.6666666666667,0.5f),
                                        new StylusPoint(83,58.6666666666667,0.5f),
                                        new StylusPoint(82.3333333333333,57.3333333333333,0.5f),
                                        new StylusPoint(81.6666666666667,56,0.5f),
                                        new StylusPoint(80.6666666666667,55,0.5f),
                                        new StylusPoint(79.6666666666667,53.6666666666667,0.5f),
                                        new StylusPoint(78.6666666666667,52.6666666666667,0.5f),
                                        new StylusPoint(77.3333333333333,51.3333333333333,0.5f),
                                        new StylusPoint(76,50.3333333333333,0.5f),
                                        new StylusPoint(74.6666666666667,49.3333333333333,0.5f),
                                        new StylusPoint(73.3333333333333,48.3333333333333,0.5f),
                                        new StylusPoint(71.6666666666667,47.3333333333333,0.5f),
                                        new StylusPoint(70,46.3333333333333,0.5f),
                                        new StylusPoint(68.3333333333333,45.6666666666667,0.5f),
                                        new StylusPoint(66.6666666666667,45,0.5f),
                                        new StylusPoint(65,44.3333333333333,0.5f),
                                        new StylusPoint(63,43.6666666666667,0.5f),
                                        new StylusPoint(61.3333333333333,43.3333333333333,0.5f),
                                        new StylusPoint(59.3333333333333,43,0.5f),
                                        new StylusPoint(57.3333333333333,42.6666666666667,0.5f),
                                        new StylusPoint(55.3333333333333,42.3333333333333,0.5f),
                                        new StylusPoint(53.3333333333333,42.3333333333333,0.5f),
                                        new StylusPoint(51,42,0.5f),
                                        new StylusPoint(49,42.3333333333333,0.5f),
                                        new StylusPoint(46.6666666666667,42.3333333333333,0.5f),
                                        new StylusPoint(44.3333333333333,42.6666666666667,0.5f),
                                        new StylusPoint(42.3333333333333,43,0.5f),
                                        new StylusPoint(40,43.6666666666667,0.5f),
                                        new StylusPoint(37.6666666666667,44.3333333333333,0.5f),
                                        new StylusPoint(35.3333333333333,45,0.5f),
                                        new StylusPoint(33,46,0.5f),
                                        new StylusPoint(30.6666666666667,47.3333333333333,0.5f),
                                        new StylusPoint(28.3333333333333,48.6666666666667,0.5f),
                                        new StylusPoint(26,50,0.5f),
                                        new StylusPoint(24,51.6666666666667,0.5f),
                                        new StylusPoint(21.6666666666667,53.6666666666667,0.5f),
                                        new StylusPoint(19.3333333333333,55.3333333333333,0.5f),
                                        new StylusPoint(17,57.3333333333333,0.5f),
                                    }
                                ),
                                (penColour!=null)?new DrawingAttributes{Color = penColour.Color, Height=2, Width=2, IsHighlighter=false}:new DrawingAttributes
                                {
                                    Color=Colors.Black,
                                    Height=2,
                                    Width=2,
                                    IsHighlighter=false
                                }
                            )
                        }
                    );
            }
            private PointCollection cachedBrushPreviewPoints;
            public PointCollection BrushPreviewPoints
            {
                get
                {
                    if (cachedBrushPreviewPoints == null)
                        generateBrushPreviewPoints();
                    return cachedBrushPreviewPoints;
                }
            }
            private void generateBrushPreviewPoints()
            {
                cachedBrushPreviewPoints = new PointCollection{
                        new Point(100,0),
                        new Point(71,0),
                        new Point(62,12),
                        new Point(62,20),
                        new Point(48,47),
                        new Point(37,65),
                        new Point(37,69),
                        new Point(31,83),
                        new Point(29,89),
                        new Point(30,90),
                        new Point(32,91),
                        new Point(37,85),
                        new Point(48,75),
                        new Point(52,75),
                        new Point(77,43),
                        new Point(91,32),
                        new Point(100,21),
                        new Point(100,0)
                    };
            }
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (pens == null || pens.Count == 0) return;
            selectPen(pens[0]);
        }
    }
    class SlideIndicator
    {
        public SlideIndicator(int Id)
        {
            slideId = Id;
        }
        public int slideId { get; private set; }
        public Slide slide
        {
            get
            {
                if (ThisAddIn.instance == null || ThisAddIn.instance.Application == null || ThisAddIn.instance.Application.ActivePresentation == null || ThisAddIn.instance.Application.ActivePresentation.Slides == null || ThisAddIn.instance.Application.ActivePresentation.Slides.Count < 1)
                    return null;
                return ThisAddIn.instance.Application.ActivePresentation.Slides.FindBySlideID(slideId);
            }
        }
        public bool isCurrentSlide { get { return slide.SlideIndex == ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition; } }
        public bool clickAdvance
        {
            get
            {
                if (slide == null) return false;
                return slide.SlideShowTransition.AdvanceOnClick == Microsoft.Office.Core.MsoTriState.msoTrue;
            }
            set
            {
                if (slide == null) return;
                slide.SlideShowTransition.AdvanceOnClick = value ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            }
        }
        public void setClickAdvance(bool state)
        {
            if (slide != null && clickAdvance != state)
            {
                slide.SlideShowTransition.AdvanceOnClick = state ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
            }
        }
    }
}