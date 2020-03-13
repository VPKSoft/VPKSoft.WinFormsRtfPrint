#region License
/*
MIT License

Copyright (c) 2020 Petteri Kautonen

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/
#endregion

using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VPKSoft.WinFormsRtfPrint
{
    // ReSharped with Re, TODO
    // (C): https://docs.microsoft.com/en-us/previous-versions/dotnet/articles/ms996492(v=msdn.10)

    /// <summary>
    /// A class for printing <see cref="RichTextBox"/> control contents.
    /// </summary>
    public class RtfPrint
    {
        #region PublicProperties        
        /// <summary>
        /// Gets or sets the <see cref="RichTextBox"/> control which contents are to be printed. The value is reset to <c>null</c> after printing.
        /// </summary>
        public static RichTextBox RichTextBox { get; set; }
        #endregion

        #region StructuresWinApi
        [StructLayout(LayoutKind.Sequential)]
        // ReSharper disable once InconsistentNaming
        private struct STRUCT_RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        // ReSharper disable once InconsistentNaming
        // ReSharper disable once IdentifierTypo
        private struct STRUCT_CHARRANGE
        {
            public int cpMin;
            public int cpMax;
        }

        [StructLayout(LayoutKind.Sequential)]
        // ReSharper disable once InconsistentNaming
        // ReSharper disable once IdentifierTypo
        private struct STRUCT_FORMATRANGE
        {
            public IntPtr hdc;
            public IntPtr hdcTarget;
            public STRUCT_RECT rc;
            public STRUCT_RECT rcPage;

            // ReSharper disable once IdentifierTypo
            public STRUCT_CHARRANGE chrg;
        }
        #endregion

        #region PInvoke
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int msg,
            int wParam, IntPtr lParam);

        // ReSharper disable once InconsistentNaming
        private const int WM_USER = 0x400;

        // ReSharper disable once InconsistentNaming
        // ReSharper disable once IdentifierTypo
        private const int EM_FORMATRANGE = WM_USER + 57;
        #endregion

        #region PrintMethods        
        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <returns><c>true</c> if no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents()
        {
            return PrintRichTextContents(null, false, false, (Icon)null, null);
        }


        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="owner">Any object that implements <see cref="T:System.Windows.Forms.IWin32Window" /> that represents the top-level window that will own the modal dialog box.</param>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(IWin32Window owner, bool showPrintPreview, bool showPrintDialog)
        {
            return PrintRichTextContents(owner, showPrintPreview, showPrintDialog, (Icon)null, null);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(bool showPrintPreview, bool showPrintDialog)
        {
            return PrintRichTextContents(null, showPrintPreview, showPrintDialog, (Icon)null, null);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="owner">Any object that implements <see cref="T:System.Windows.Forms.IWin32Window" /> that represents the top-level window that will own the modal dialog box.</param>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="assembly">An assembly to get an icon for the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(IWin32Window owner, bool showPrintPreview, bool showPrintDialog,
            Assembly assembly, string previewDialogTitle)
        {
            var icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);
            return PrintRichTextContents(owner, showPrintPreview, showPrintDialog, icon, previewDialogTitle);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="assembly">An assembly to get an icon for the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(bool showPrintPreview, bool showPrintDialog,
            Assembly assembly, string previewDialogTitle)
        {
            var icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);
            return PrintRichTextContents(null, showPrintPreview, showPrintDialog, icon, previewDialogTitle);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="owner">Any object that implements <see cref="T:System.Windows.Forms.IWin32Window" /> that represents the top-level window that will own the modal dialog box.</param>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="iconForm">A <see cref="Form"/> class instance to get the icon to be used with the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(IWin32Window owner, bool showPrintPreview, bool showPrintDialog,
            Form iconForm, string previewDialogTitle)
        {
            var icon = iconForm.Icon;
            return PrintRichTextContents(owner, showPrintPreview, showPrintDialog, icon, previewDialogTitle);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="iconForm">A <see cref="Form"/> class instance to get the icon to be used with the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(bool showPrintPreview, bool showPrintDialog,
            Form iconForm, string previewDialogTitle)
        {
            var icon = iconForm.Icon;
            return PrintRichTextContents(null, showPrintPreview, showPrintDialog, icon, previewDialogTitle);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="icon">The icon to use with the <see cref="PrintPreviewDialog"/> instance.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(bool showPrintPreview, bool showPrintDialog, Icon icon, string previewDialogTitle)
        {
            return PrintRichTextContents(null, showPrintPreview, showPrintDialog, icon, previewDialogTitle);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="owner">Any object that implements <see cref="T:System.Windows.Forms.IWin32Window" /> that represents the top-level window that will own the modal dialog box.</param>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="assembly">An assembly to get an icon for the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(IWin32Window owner, bool showPrintPreview, bool showPrintDialog,
            Assembly assembly)
        {
            var icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);
            return PrintRichTextContents(owner, showPrintPreview, showPrintDialog, icon, null);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="assembly">An assembly to get an icon for the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(bool showPrintPreview, bool showPrintDialog,
            Assembly assembly)
        {
            var icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);
            return PrintRichTextContents(null, showPrintPreview, showPrintDialog, icon, null);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="owner">Any object that implements <see cref="T:System.Windows.Forms.IWin32Window" /> that represents the top-level window that will own the modal dialog box.</param>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="iconForm">A <see cref="Form"/> class instance to get the icon to be used with the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(IWin32Window owner, bool showPrintPreview, bool showPrintDialog,
            Form iconForm)
        {
            var icon = iconForm.Icon;
            return PrintRichTextContents(owner, showPrintPreview, showPrintDialog, icon, null);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="iconForm">A <see cref="Form"/> class instance to get the icon to be used with the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(bool showPrintPreview, bool showPrintDialog,
            Form iconForm)
        {
            var icon = iconForm.Icon;
            return PrintRichTextContents(null, showPrintPreview, showPrintDialog, icon, null);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="icon">The icon to use with the <see cref="PrintPreviewDialog"/> instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(bool showPrintPreview, bool showPrintDialog, Icon icon)
        {
            return PrintRichTextContents(null, showPrintPreview, showPrintDialog, icon, null);
        }

        /// <summary>
        /// Prints the contents of the <see cref="RichTextBox"/> instance.
        /// </summary>
        /// <param name="owner">Any object that implements <see cref="T:System.Windows.Forms.IWin32Window" /> that represents the top-level window that will own the modal dialog box.</param>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="icon">The icon to use with the <see cref="PrintPreviewDialog"/> instance.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintRichTextContents(IWin32Window owner, bool showPrintPreview, bool showPrintDialog, Icon icon, string previewDialogTitle)
        {
            try
            {
                using (var printDoc = new PrintDocument())
                {
                    if (showPrintDialog)
                    {
                        using (var printDialog = new PrintDialog())
                        {
                            if (printDialog.ShowDialog(owner) == DialogResult.OK)
                            {
                                printDoc.PrinterSettings = printDialog.PrinterSettings;
                            }
                            else
                            {
                                RichTextBox = null;
                                return false;
                            }
                        }
                    }


                    printDoc.BeginPrint += printDoc_BeginPrint;
                    printDoc.PrintPage += printDoc_PrintPage;
                    printDoc.EndPrint += printDoc_EndPrint;

                    if (showPrintPreview)
                    {
                        using (var pdDialog = new PrintPreviewDialog())
                        {
                            pdDialog.Document = printDoc;
                            if (icon != null)
                            {
                                pdDialog.Icon = icon;
                            }

                            pdDialog.WindowState = FormWindowState.Maximized;
                            if (previewDialogTitle != null)
                            {
                                pdDialog.Text = previewDialogTitle;
                            }

                            if (pdDialog.ShowDialog(owner) != DialogResult.OK)
                            {
                                RichTextBox = null;
                                return false;
                            }
                        }
                    }

                    // Start printing process
                    printDoc.Print();

                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region HelperMethods
        /// <summary>
        /// Convert between 1/100 inch (unit used by the .NET framework)
        /// and twips (1/1440 inch, used by Win32 API calls)
        /// </summary>
        /// <param name="n">Value in 1/100 inch</param>
        /// <returns>Value in twips</returns>
        private static int HundredthInchToTwips(int n)
        {
            return (int) (n * 14.4);
        }

        /// <summary>
        /// Free cached data from rich edit control after printing
        /// </summary>
        /// <param name="handle">The handle to the <see cref="RichTextBox"/> control</param>
        private static void FormatRangeDone(IntPtr handle)
        {
            IntPtr lParam = new IntPtr(0);
            SendMessage(handle, EM_FORMATRANGE, 0, lParam);
        }

        /// <summary>
        /// Calculate or render the contents of our RichTextBox for printing
        /// </summary>
        /// <param name="measureOnly">If true, only the calculation is performed,
        /// otherwise the text is rendered as well</param>
        /// <param name="e">The PrintPageEventArgs object from the
        /// PrintPage event</param>
        /// <param name="charFrom">Index of first character to be printed</param>
        /// <param name="charTo">Index of last character to be printed</param>
        /// <param name="handle">The handle to the <see cref="RichTextBox"/> control</param>
        /// <returns>(Index of last character that fitted on the
        /// page) + 1</returns>
        private static int FormatRange(IntPtr handle, bool measureOnly, PrintPageEventArgs e,
            int charFrom, int charTo)
        {
            // Specify which characters to print
            STRUCT_CHARRANGE cr;
            cr.cpMin = charFrom;
            cr.cpMax = charTo;

            // Specify the area inside page margins
            STRUCT_RECT rc;
            rc.top = HundredthInchToTwips(e.MarginBounds.Top);
            rc.bottom = HundredthInchToTwips(e.MarginBounds.Bottom);
            rc.left = HundredthInchToTwips(e.MarginBounds.Left);
            rc.right = HundredthInchToTwips(e.MarginBounds.Right);

            // Specify the page area
            STRUCT_RECT rcPage;
            rcPage.top = HundredthInchToTwips(e.PageBounds.Top);
            rcPage.bottom = HundredthInchToTwips(e.PageBounds.Bottom);
            rcPage.left = HundredthInchToTwips(e.PageBounds.Left);
            rcPage.right = HundredthInchToTwips(e.PageBounds.Right);

            // Get device context of output device
            IntPtr hdc = e.Graphics.GetHdc();

            // ReSharper disable once CommentTypo
            // Fill in the FORMATRANGE struct
            STRUCT_FORMATRANGE fr;
            fr.chrg = cr;
            fr.hdc = hdc;
            fr.hdcTarget = hdc;
            fr.rc = rc;
            fr.rcPage = rcPage;

            // Non-Zero wParam means render, Zero means measure
            int wParam = (measureOnly ? 0 : 1);

            // ReSharper disable once CommentTypo
            // Allocate memory for the FORMATRANGE struct and
            // copy the contents of our struct to this memory
            IntPtr lParam = Marshal.AllocCoTaskMem(Marshal.SizeOf(fr));
            Marshal.StructureToPtr(fr, lParam, false);

            // Send the actual Win32 message
            int res = SendMessage(handle, EM_FORMATRANGE, wParam, lParam);

            // Free allocated memory
            Marshal.FreeCoTaskMem(lParam);

            // and release the device context
            e.Graphics.ReleaseHdc(hdc);

            return res;
        }

        #endregion

        #region InternalEvents
        // variable to trace text to print for pagination
        private static int _mNFirstCharOnPage;

        private static void printDoc_BeginPrint(object sender,
            PrintEventArgs e)
        {
            // Start at the beginning of the text
            _mNFirstCharOnPage = 0;
        }

        private static void printDoc_PrintPage(object sender,
            PrintPageEventArgs e)
        {
            // To print the boundaries of the current page margins
            // uncomment the next line:
            // e.Graphics.DrawRectangle(System.Drawing.Pens.Blue, e.MarginBounds);

            // make the RichTextBoxEx calculate and render as much text as will
            // fit on the page and remember the last character printed for the
            // beginning of the next page
            _mNFirstCharOnPage = FormatRange(RichTextBox.Handle, false,
                e,
                _mNFirstCharOnPage,
                RichTextBox.TextLength);

            // check if there are more pages to print
            e.HasMorePages = _mNFirstCharOnPage < RichTextBox.TextLength;
        }

        private static void printDoc_EndPrint(object sender,
            PrintEventArgs e)
        {
            // Clean up cached information
            FormatRangeDone(RichTextBox.Handle);
            RichTextBox = null;
        }
        #endregion
    }
}
