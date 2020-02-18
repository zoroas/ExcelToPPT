using System;
using System.Linq;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace ExcelToPPT
{
    public class PowerPointCreator
    {
        private bool hasError = true;
        private string error = "";

        public bool HasError()
        {
            return this.hasError;
        }

        public string GetError()
        {
            return this.error;
        }

        public static void InsertImageInSlide(PowerPoint.Slide slide, PowerPoint.Shape shape, string imagePath)
        {
            float left = shape.Left;
            float top = shape.Top;
            float width = shape.Width;
            float height = shape.Height;
            slide.Shapes.AddPicture(imagePath,Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,left, top, width, height);
        }

        public PowerPointCreator(
                String excelFile, string pptTemplate, string pptOutput, 
                string rowNumberToGetColumnsStr, string numberOfColumnsStr, String photoFolder)
        {
            PowerPoint.Application oPowerPoint = null;
            PowerPoint.Presentations oPres = null;
            PowerPoint.Presentation oPre = null;
            PowerPoint.Slides oSlides = null;
            PowerPoint.Slide oSlide = null;
            PowerPoint.Shapes oShapes = null;
            PowerPoint.Shape oShape = null;
            PowerPoint.TextFrame oTxtFrame = null;
            PowerPoint.TextRange oTxtRange = null;

            Excel.Application oExcel = null;
            Excel.Workbook oWorkBook = null;
            Excel.Worksheet oSheet = null;

            try
            {
                int rowNumberToGetColumns = int.Parse(rowNumberToGetColumnsStr);
                int numberOfColumns = int.Parse(numberOfColumnsStr);

                // Create the PowerPoint
                oPowerPoint = new PowerPoint.Application( );
                oPres = oPowerPoint.Presentations;
                File.Copy(pptTemplate, pptOutput);

                oPre = oPres.Open(pptOutput);
                oSlides = oPre.Slides;

                // Open Excel
                oExcel = new Excel.Application( );
                oWorkBook = oExcel.Workbooks.Open(excelFile);
                oSheet = oWorkBook.Sheets[1];

                // Iterate over lines
                bool doMore = true;
                int row = rowNumberToGetColumns;

                while (doMore){

                    row++;

                    oSlide = oSlides[1];
                    oSlide.Duplicate( );
                    oSlide = oSlides[2];
                    doMore = false;
                    for (int col = 1; col <= numberOfColumns; col++)
                    {
                        Excel.Range label = oSheet.Cells[rowNumberToGetColumns, col];
                        String myLabel = (":"+label.Value + ":").ToLower();
                        Excel.Range cell = oSheet.Cells[row, col];

                        String myValue = "";

                        try
                        {
                            if (cell.Value != null)
                            {
                                DateTime date = cell.Value;
                                myValue = date.ToString(this.GetDateFormat());
                            }
                        }
                        catch (Exception)
                        {
                            try
                            {
                                myValue = cell.Value.ToString();
                            }
                            catch (Exception)
                            {
                            }
                        }

                        if (myValue != "")
                        {
                            doMore = true;
                        }

                        if (myLabel == ":photo:")
                        {
                            myValue = photoFolder + "/" + myValue + ".jpg";
                            for (int i = 1; i <= oSlide.Shapes.Count; i++)
                            {
                                PowerPoint.Shape shape = oSlide.Shapes[i];
                                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                {
                                    String shapeLabel = shape.TextFrame.TextRange.Text;
                                    if (shapeLabel.ToLower() == ":photo:" && File.Exists(myValue))
                                    {
                                        float left = shape.Left;
                                        float top = shape.Top;
                                        float width = shape.Width;
                                        float height = shape.Height;
                                        oSlide.Shapes.AddPicture(myValue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, left, top, width, height);
                                        shape.Delete();
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int i = 1; i <= oSlide.Shapes.Count; i++)
                            {
                                PowerPoint.Shape shape = oSlide.Shapes[i];
                                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                {
                                    String oldValue = shape.TextFrame.TextRange.Text;
                                    int first = oldValue.ToLower().IndexOf(myLabel);
                                    if (first != -1)
                                    {
                                        String newValue = oldValue.Substring(0,first != 0 ? first : 0) + 
                                                          myValue +              
                                                          oldValue.Substring(first + myLabel.Length);
                                        if (oldValue != newValue)
                                        {
                                            shape.TextFrame.TextRange.Text = newValue;
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
                oSlides[1].Delete( );
                oPowerPoint.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                //foreach()
                //oShapes = oSlide.Shapes;
                //oShape = oShapes[1];
                //oTxtFrame = oShape.TextFrame;
                //oTxtRange = oTxtFrame.TextRange;
                //oTxtRange.Text = "All-In-One Code Framework";

                //// Save the presentation as a pptx file and close it.

                //Console.WriteLine("Save and close the presentation");

                //string fileName = Path.GetDirectoryName(
                //    Assembly.GetExecutingAssembly().Location) + "\\Sample1.pptx";
                //oPre.SaveAs(fileName,
                //    PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                //    Office.MsoTriState.msoTriStateMixed);
                //oPre.Close( );

                //// Quit the PowerPoint application.

                //Console.WriteLine("Quit the PowerPoint application");
                //oPowerPoint.Quit( );
                this.hasError = false;
            }
            catch (Exception ex)
            {
                this.error = ex.Message;
            }
            finally
            {
                // Clean up the unmanaged PowerPoint COM resources by explicitly 
                // calling Marshal.FinalReleaseComObject on all accessor objects. 
                // See http://support.microsoft.com/kb/317109.

                if (oTxtRange != null)
                {
                    Marshal.FinalReleaseComObject(oTxtRange);
                    oTxtRange = null;
                }
                if (oTxtFrame != null)
                {
                    Marshal.FinalReleaseComObject(oTxtFrame);
                    oTxtFrame = null;
                }
                if (oShape != null)
                {
                    Marshal.FinalReleaseComObject(oShape);
                    oShape = null;
                }
                if (oShapes != null)
                {
                    Marshal.FinalReleaseComObject(oShapes);
                    oShapes = null;
                }
                if (oSlide != null)
                {
                    Marshal.FinalReleaseComObject(oSlide);
                    oSlide = null;
                }
                if (oSlides != null)
                {
                    Marshal.FinalReleaseComObject(oSlides);
                    oSlides = null;
                }
                if (oPre != null)
                {
                    Marshal.FinalReleaseComObject(oPre);
                    oPre = null;
                }
                if (oPres != null)
                {
                    Marshal.FinalReleaseComObject(oPres);
                    oPres = null;
                }
                if (oPowerPoint != null)
                {
                    Marshal.FinalReleaseComObject(oPowerPoint);
                    oPowerPoint = null;
                }
                if (oSheet != null)
                {
                    Marshal.FinalReleaseComObject(oSheet);
                    oSheet = null;
                }
                if (oWorkBook != null)
                {
                    oWorkBook.Close( );
                    Marshal.FinalReleaseComObject(oWorkBook);
                    oWorkBook = null;
                }
            }
        }

        private string GetDateFormat()
        {
            string format = MySettings.Default.SettingDateFormat;
            if (String.IsNullOrEmpty(format))
            {
                format = DateTimeFormatInfo.CurrentInfo.GetAllDateTimePatterns('d').First();
                MySettings.Default.SettingDateFormat = format;
                MySettings.Default.Save();
            }
            return format;
        }

    }
}
