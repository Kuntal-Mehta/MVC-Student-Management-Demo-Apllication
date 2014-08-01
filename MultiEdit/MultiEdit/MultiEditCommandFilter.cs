using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Text.Editor;
using System.Windows.Input;
using System.Collections.Generic;
using Microsoft.VisualStudio.Text;
using System.Windows.Media;
using System.Windows.Controls;

namespace MultiEdit
{
    class MultiEditCommandFilter : IOleCommandTarget
    {
        private IWpfTextView m_textView;
        internal IOleCommandTarget m_nextTarget;
        internal bool m_added;
        internal bool m_adorned;
        private IAdornmentLayer m_adornmentLayer;
        private bool requiresHandling = false;
        List<ITrackingPoint> trackList = new List<ITrackingPoint>();
        public MultiEditCommandFilter(IWpfTextView textView)
        {
            m_textView = textView;
            m_adorned = false;
            m_adornmentLayer = m_textView.GetAdornmentLayer("MultiEditLayer");
            m_textView.LayoutChanged += m_textView_LayoutChanged;
        }

        int IOleCommandTarget.QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return m_nextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        int IOleCommandTarget.Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            requiresHandling = false;
            // When Alt Clicking, we need to add Edit points.
            if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.ECMD_LEFTCLICK && Keyboard.Modifiers == ModifierKeys.Alt)
            {
                requiresHandling = true;
            }
            else if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.ECMD_LEFTCLICK && trackList.Count > 0)
            {
                requiresHandling = true;
            }
            else if (pguidCmdGroup == VSConstants.VSStd2K && trackList.Count > 0 && (nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR ||
                    nCmdID == (uint)VSConstants.VSStd2KCmdID.BACKSPACE ||
                    nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB ||
                    nCmdID == (uint)VSConstants.VSStd2KCmdID.UP ||
                    nCmdID == (uint)VSConstants.VSStd2KCmdID.DOWN ||
                    nCmdID == (uint)VSConstants.VSStd2KCmdID.LEFT ||
                    nCmdID == (uint)VSConstants.VSStd2KCmdID.RIGHT
            ))
            {
                requiresHandling = true;  
            }
            if (requiresHandling == true)
            {
                // Capture Alt Left Click, only when the Box Selection mode hasn't been used (After Drag-selecting)
                if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.ECMD_LEFTCLICK &&
                                                            Keyboard.Modifiers == ModifierKeys.Alt)
                {
                    // Add a Edit point, show it Visually 
                    AddSyncPoint();
                    RedrawScreen();
                }
                else if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.ECMD_LEFTCLICK &&
                 trackList.Count > 0)
                {
                    // Switch back to normal, clear out Edit points
                    ClearSyncPoints();
                    RedrawScreen();
                }
                else if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR)
                {
                    var typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);
                    InsertSyncedChar(typedChar.ToString());
                    RedrawScreen();
                }
                return m_nextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }
            return m_nextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }
        private void InsertSyncedChar(string inputString)
        {
            // Avoiding inserting the character for the last edit point, as the Caret is there and
            // the default IDE behavior will insert the text as expected.
            ITextEdit edit = m_textView.TextBuffer.CreateEdit();
            for (int i = 0; i < trackList.Count - 1; i++)
            {
                var curTrackPoint = trackList[i];
                edit.Insert(curTrackPoint.GetPosition(m_textView.TextSnapshot), inputString);
            }
            edit.Apply();
            edit.Dispose();
        }
        private void SyncedBackSpace()
        {
            ITextEdit edit = m_textView.TextBuffer.CreateEdit();
            for (int i = 0; i < trackList.Count - 1; i++)
            {
                var curTrackPoint = trackList[i];
                edit.Delete(curTrackPoint.GetPosition(m_textView.TextSnapshot) - 1, 1);
            }
            edit.Apply();
            edit.Dispose();
        }
        private void SyncedDelete()
        {
            ITextEdit edit = m_textView.TextBuffer.CreateEdit();
            for (int i = 0; i < trackList.Count - 1; i++)
            {
                var curTrackPoint = trackList[i];
                edit.Delete(curTrackPoint.GetPosition(m_textView.TextSnapshot), 1);
            }
            edit.Apply();
            edit.Dispose();
        }
        private void ClearSyncPoints()
        {
            trackList.Clear();
            m_adornmentLayer.RemoveAllAdornments();
        }
        private void RedrawScreen()
        {
            m_adornmentLayer.RemoveAllAdornments();
            for (int i = 0; i < trackList.Count; i++)
            {
                var curTrackPoint = trackList[i];
                DrawSingleSyncPoint(curTrackPoint);
            }
        }
        private void AddSyncPoint()
        {
            // Get the Caret location, and Track it
            CaretPosition curPosition = m_textView.Caret.Position;
            var curTrackPoint = m_textView.TextSnapshot.CreateTrackingPoint(curPosition.BufferPosition.Position,
            Microsoft.VisualStudio.Text.PointTrackingMode.Positive);
            trackList.Add(curTrackPoint);
        }
        private void DrawSingleSyncPoint(ITrackingPoint curTrackPoint)
        {
            SnapshotSpan span;
            span = new SnapshotSpan(curTrackPoint.GetPoint(m_textView.TextSnapshot), 1);

            var brush = new SolidColorBrush(Colors.LightPink);
            var g = m_textView.TextViewLines.GetLineMarkerGeometry(span);
            GeometryDrawing drawing = new GeometryDrawing(brush, null, g);
            if (drawing.Bounds.IsEmpty)
                return;

            System.Windows.Shapes.Rectangle r = new System.Windows.Shapes.Rectangle()
            {
                Fill = brush,
                Width = drawing.Bounds.Width / 2,
                Height = drawing.Bounds.Height
            };
            Canvas.SetLeft(r, g.Bounds.Left);
            Canvas.SetTop(r, g.Bounds.Top);
            m_adornmentLayer.AddAdornment(AdornmentPositioningBehavior.TextRelative,
                                         span, "MultiEditLayer", r, null);
        }
        private void m_textView_LayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            RedrawScreen();
        }

    }
}