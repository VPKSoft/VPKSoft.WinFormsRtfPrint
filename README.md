# VPKSoft.WinFormsRtfPrint
A library to help to print the contents of a [RichTextBox](https://docs.microsoft.com/en-us/dotnet/framework/winforms/controls/richtextbox-control-overview-windows-forms) control.

## Sample
```cs
using VPKSoft.WinFormsRtfPrint;

namespace TestApp
{
    public partial class FormMain : Form
    {
        private void mnuPrint_Click(object sender, EventArgs e)
        {
            rtbPrintTest.PrintWithDialog();
        }
    }
}
```
**_OR_**

Use the static method overloads of the RtfPrint class:
```
using VPKSoft.WinFormsRtfPrint;

namespace TestApp
{
    public partial class FormMain : Form
    {
        private void mnuPrint_Click(object sender, EventArgs e)
        {
            RtfPrint.RichTextBox = rtbPrintTest;
            RtfPrint.PrintRichTextContents();
        }
    }
}
```

This library is based on the [Microsoft article](https://docs.microsoft.com/en-us/previous-versions/dotnet/articles/ms996492(v=msdn.10))  [Getting WYSIWYG Print Results from a .NET RichTextBox](https://docs.microsoft.com/en-us/previous-versions/dotnet/articles/ms996492(v=msdn.10)) by Martin Müller.

## Thanks to
* [JetBrains](http://www.jetbrains.com) for their open source license(s).

![JetBrains](http://www.vpksoft.net/site/External/JetBrains/jetbrains.svg)
