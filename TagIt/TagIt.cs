using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace TagIt
{
       
    class TagIt
    {
        private IWpfTextView _view;
        private IAdornmentLayer _layer;
        private DTE2 _dte;
        List<ITrackingPoint> trackingPointList = new List<ITrackingPoint>();

        public TagIt(IWpfTextView view, DTE2 dte)
        {
            _view = view;
            _layer = _view.GetAdornmentLayer("TagIt");
            _dte = dte;
            //Currently I dont have specific idea at whcih event should we trigger this thing. It is subject to analysis.
            _view.LayoutChanged += _view_LayoutChanged;
        }

        void _view_LayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            LayoutChange();
        }

        private void LayoutChange()
        {
            if(this._dte!=null && this._dte.ActiveDocument!=null)
            {
                List<CodeClass> foundClasses = new List<CodeClass>();
                List<CodeFunction> foundMethod = new List<CodeFunction>();
                RecursiveClassSearch(this._dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements, foundClasses);
                RecursiveMethodSearch(this._dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements, foundMethod);
                GettingToTheLocationOfInsertionOfClass(foundClasses);
                GettingToTheLocationOfInsertionOfFunction(foundMethod);
                redrawScreen();
            }
        }

        private void ClearScreen()
        {
            trackingPointList.Clear();
            _layer.RemoveAllAdornments();
        }
       
        private static void RecursiveClassSearch(CodeElements elements, List<CodeClass> foundClasses)
        {
            foreach (CodeElement codeElement in elements)
            {
                if (codeElement is CodeClass)
                {
                    foundClasses.Add(codeElement as CodeClass);
                }
                RecursiveClassSearch(codeElement.Children, foundClasses);
            }
        }

        public static void RecursiveMethodSearch(CodeElements elements, List<CodeFunction> foundMethod)
        {
            foreach (CodeElement codeElement in elements)
            {
                if(codeElement is CodeFunction)
                {
                    foundMethod.Add(codeElement as CodeFunction);
                }
                RecursiveMethodSearch(codeElement.Children, foundMethod);
            }
        }

        private void GettingToTheLocationOfInsertionOfClass(List<CodeClass> foundClasses)
        {
            foreach (CodeElement Class in foundClasses)
            {     
                Editing(Class);
            }
        }

        private void GettingToTheLocationOfInsertionOfFunction(List<CodeFunction> foundMethod)
        {
            foreach (CodeElement method in foundMethod)
            {             
                Editing(method);
            }
        }

        private void Editing(CodeElement element)
        {
            #region Commented Demo
            //Brush brush = new SolidColorBrush(Colors.BlueViolet);
            //brush.Freeze();
            //Brush penBrush = new SolidColorBrush(Colors.Red);
            //penBrush.Freeze();
            //Pen pen = new Pen(penBrush, 0.5);
            //pen.Freeze();

            ////draw a square with the created brush and pen
            //System.Windows.Rect r = new System.Windows.Rect(10, 10, 8, 8);
            //Geometry g = new RectangleGeometry(r);
            //GeometryDrawing drawing = new GeometryDrawing(brush, pen, g);
            //drawing.Freeze();

            //DrawingImage drawingImage = new DrawingImage(drawing);
            //drawingImage.Freeze(); 
            #endregion
            #region Commented Demo Display
            ////Some drawing methods will go here. Now it depends upon how to show the code.
            //var linenumber = element.StartPoint.Line;
            //var offset = element.StartPoint.LineCharOffset;
            //TextSelection selection = _dte.ActiveDocument.Selection as TextSelection;
            //selection.MoveToLineAndOffset(linenumber, offset, false);
            //selection.Insert("Demo "); 
            #endregion
            addTrackingPoint(element);
        }

        private void redrawScreen()
        {
            _layer.RemoveAllAdornments();
            for (int i=0;i<trackingPointList.Count;i++)
            {
                var curTrackPoint = trackingPointList[i];
                CreateLabel(curTrackPoint);
            }
        }

        private void CreateLabel(ITrackingPoint curTrackPoint)
        {
            SnapshotSpan span;
            span = new SnapshotSpan(curTrackPoint.GetPoint(_view.TextSnapshot), 1);
            var brush = new SolidColorBrush(Colors.LightPink);
            var geometry = _view.TextViewLines.GetLineMarkerGeometry(span);
            GeometryDrawing drawing = new GeometryDrawing(brush, null, geometry);
            if (drawing.Bounds.IsEmpty)
                return;
            Rectangle rectangle =new Rectangle()
            {
                Fill = brush,
                Width = drawing.Bounds.Width,
                Height = drawing.Bounds.Height,
            };
            Canvas.SetLeft(rectangle, geometry.Bounds.X);
            Canvas.SetTop(rectangle, geometry.Bounds.Y);
            _layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, "TagIt", rectangle, null);
        }

        private void addTrackingPoint(CodeElement element)
        {
            TextPoint startPoint = element.StartPoint;
            int CaretPosition = ConvertToPosition(startPoint);
            var curTrackPoint = _view.TextSnapshot.CreateTrackingPoint(CaretPosition, PointTrackingMode.Positive);
            trackingPointList.Add(curTrackPoint);
        }

        private int ConvertToPosition(TextPoint editPoint)
        {
            return editPoint.AbsoluteCharOffset + editPoint.Line-2;
        }
    }
}