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

using System.Drawing;
using System.Windows.Forms;

namespace VPKSoft.WinFormsRtfPrint
{
    /// <summary>
    /// A helper class to print the contents of a <see cref="RichTextBox"/> control contents.
    /// </summary>
    public static class RichTextBoxPrintHelper
    {
        /// <summary>
        /// Prints the document.
        /// </summary>
        /// <param name="richTextBox"></param>
        /// <returns><c>true</c> if the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool Print(this RichTextBox richTextBox)
        {
            try
            {
                // IWin32Window owner, bool showPrintPreview, bool showPrintDialog, Icon icon, string previewDialogTitle
                RtfPrint.RichTextBox = richTextBox;
                return RtfPrint.PrintRichTextContents(null, false, false, (Icon)null, null);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Prints the document optionally displaying a <see cref="PrintDialog"/> and a <see cref="PrintPreviewDialog"/> to the user before printing.
        /// </summary>
        /// <param name="richTextBox">The <see cref="RichTextBox"/> class instance which contents are to be printed.</param>
        /// <param name="showPrintPreview">if set to <c>true</c> a print preview dialog is shown before printing the document.</param>
        /// <param name="showPrintDialog">if set to <c>true</c> a <see cref="PrintDialog"/> is shown before printing.</param>
        /// <param name="icon">A <see cref="Form"/> class instance to get the icon to be used with the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool Print(this RichTextBox richTextBox, bool showPrintPreview, bool showPrintDialog, Icon icon, string previewDialogTitle)
        {
            try
            {
                // IWin32Window owner, bool showPrintPreview, bool showPrintDialog, Icon icon, string previewDialogTitle
                RtfPrint.RichTextBox = richTextBox;
                return RtfPrint.PrintRichTextContents(null, showPrintPreview, showPrintDialog, icon, previewDialogTitle);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Only shows the print preview dialog without printing the document.
        /// </summary>
        /// <param name="richTextBox">The <see cref="RichTextBox"/> class instance which contents are to be previewed.</param>
        /// <param name="icon">A <see cref="Form"/> class instance to get the icon to be used with the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintPreview(this RichTextBox richTextBox, Icon icon, string previewDialogTitle)
        {
            try
            {
                // IWin32Window owner, bool showPrintPreview, bool showPrintDialog, Icon icon, string previewDialogTitle
                RtfPrint.RichTextBox = richTextBox;
                return RtfPrint.PrintRichTextContents(null, true, false, icon, previewDialogTitle, true);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Prints the document displaying a <see cref="PrintDialog"/> to the user before printing.
        /// </summary>
        /// <param name="richTextBox">The <see cref="RichTextBox"/> class instance which contents are to be printed.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintWithDialog(this RichTextBox richTextBox)
        {
            try
            {
                // IWin32Window owner, bool showPrintPreview, bool showPrintDialog, Icon icon, string previewDialogTitle
                RtfPrint.RichTextBox = richTextBox;
                return RtfPrint.PrintRichTextContents(null, false, true, (Icon)null, null);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Prints the document displaying a <see cref="PrintPreviewDialog"/> to the user before printing.
        /// </summary>
        /// <param name="richTextBox">The <see cref="RichTextBox"/> class instance which contents are to be printed.</param>
        /// <param name="icon">A <see cref="Form"/> class instance to get the icon to be used with the <see cref="PrintPreviewDialog"/> dialog.</param>
        /// <param name="previewDialogTitle">The title to use with the <see cref="PrintPreviewDialog"/> class instance.</param>
        /// <returns><c>true</c> if the user accepted the optional dialogs, no exceptions were thrown and the document was printed successfully, <c>false</c> otherwise.</returns>
        public static bool PrintWithPreview(this RichTextBox richTextBox, Icon icon, string previewDialogTitle)
        {
            try
            {
                // IWin32Window owner, bool showPrintPreview, bool showPrintDialog, Icon icon, string previewDialogTitle
                RtfPrint.RichTextBox = richTextBox;
                return RtfPrint.PrintRichTextContents(null, false, true, icon, previewDialogTitle);
            }
            catch
            {
                return false;
            }
        }

    }
}
