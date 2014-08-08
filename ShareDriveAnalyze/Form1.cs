using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using ZedGraph;

namespace alphaDrive
{
    
    
    public partial class Form1 : Form
    {
        private string DrivePath = "c:\\";
        delegate void SetTextCallback(string text);
        delegate void UpdateProcessSizeDelegate(int value);
        delegate void UpdateProgressStep(int vlaue);
        delegate void EnableResultsButton();
        string currenResultsFile = "";
        
        public Form1()
        {
            InitializeComponent();
            if (!new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "results\\").Exists)
                new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "results\\").Create();
            if (new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "results\\").GetFiles().Length > 0)
                this.button5.Image = this.imageList3.Images[1];
            
        }

        public void setDrivePath(string newPath)
        {
            this.DrivePath = newPath;
        }

        public string getDrivePath()
        {
           return this.DrivePath;
        }

        public void setStatusMsg(string newMessage)
        {
            this.doingLabel.Text = newMessage;
            this.Refresh();
        }

        public void setFileMsg(string newMessage)
        {
            this.fileLabel.Text = newMessage;
            this.Refresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.setDrivePath(this.folderBrowserDialog1.SelectedPath);
                this.pictureBox1.Image = this.imageList2.Images[0];
                this.button2.Image = this.imageList1.Images[0];
                this.button2.Enabled = true;
                this.setStatusMsg("We have a job to do");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.button1.Enabled = false;
            this.button2.Enabled = false;
            this.button3.Enabled = false;
            Thread processingThread = new Thread(new ThreadStart(this.processDrive));
            processingThread.Start();
        }

        void UpdateProgressSize(int value)
        {
            this.progressBar1.Maximum = value;
            this.progressBar1.Minimum = 0;
        }

        void UpdateProcessStep(int value)
        {
            this.progressBar1.Value = value;
        }

        void EnableResultsBtn()
        {
            this.pictureBox2.Image = this.imageList2.Images[0];
            this.button3.Image = this.imageList1.Images[1];
            this.button5.Image = this.imageList3.Images[1];
            this.button3.Enabled = true;
            this.setStatusMsg("ALL DONE!  That was hard work");
        }

        private void processDrive()
        {
            DateTime start = DateTime.Now;
            DateTime end;
            DateTime oldestFileDate = DateTime.Now;
            string oldestFileName = "";
            float totalSize = 0;
            
            SetTextCallback d = new SetTextCallback(setStatusMsg);
            this.Invoke(d, new object[] { "Finding out stuff about the drive" });

            d = new SetTextCallback(setStatusMsg);
            this.Invoke(d, new object[] { "Building list of files" });
            object[] lists = GetFilesRecursive.GetFiles(this.folderBrowserDialog1.SelectedPath);

            DirectoryInfo root = new DirectoryInfo(this.folderBrowserDialog1.SelectedPath);

            List<string> allFiles = (List<string>)lists[0];
            List<string> allDirs = (List<string>)lists[1];

            allDirs.Sort();
            //remove duplicates
            Int32 index = 0;
            while (index < allDirs.Count - 1)
            {
                if (allDirs[index] == allDirs[index + 1])
                    allDirs.RemoveAt(index);
                else
                    index++;

            }

            string taxonomoy = this.BuildTaxonomy(allDirs, this.folderBrowserDialog1.SelectedPath);

            GC.WaitForPendingFinalizers();
            GC.Collect();

            int progressSize = allFiles.Count/10;

            UpdateProcessSizeDelegate updateProcessSize = new UpdateProcessSizeDelegate(UpdateProgressSize);
            UpdateProgressStep updateProgressStep = new UpdateProgressStep(UpdateProcessStep);

            this.progressBar1.Invoke(updateProcessSize, new object[] { progressSize });

            d = new SetTextCallback(setStatusMsg);
            this.Invoke(d, new object[] { "Finding out stuff about those files" });

            List<string> formats = new List<string>();
            List<int> formatInstance = new List<int>();
            List<float> formatSize = new List<float>();
            List<string> fileAuthors = new List<string>();
            List<int> authorInstance = new List<int>();
            List<string> fileAppliecation = new List<string>();
            List<int> applicationInstance = new List<int>();

            DSOFile.OleDocumentPropertiesClass dsoFile = new DSOFile.OleDocumentPropertiesClass();
            
            for (int i = 0; i < allFiles.Count; i++)
            {
                d = new SetTextCallback(setFileMsg);
                this.Invoke(d, new object[] { allFiles[i] });

                FileInfo currentFile = new FileInfo(allFiles[i]);

                if (currentFile.CreationTime.CompareTo(oldestFileDate) < 0)
                {
                    oldestFileDate = currentFile.CreationTime;
                    oldestFileName = currentFile.Name;
                }

                try 
                { 
                    dsoFile.Open(currentFile.FullName, true, DSOFile.dsoFileOpenOptions.dsoOptionDefault);

                    if (dsoFile.SummaryProperties.Author != null)
                    {
                        if (!fileAuthors.Contains(dsoFile.SummaryProperties.Author))
                        {
                            fileAuthors.Add(dsoFile.SummaryProperties.Author);
                            authorInstance.Add(1);
                        }
                        else if (fileAuthors.Contains(dsoFile.SummaryProperties.Author))
                            authorInstance[fileAuthors.IndexOf(dsoFile.SummaryProperties.Author)] = authorInstance[fileAuthors.IndexOf(dsoFile.SummaryProperties.Author)] + 1;
                    }

                }
                catch (Exception ec) { };

                if (!formats.Contains(currentFile.Extension) && currentFile.Exists)
                {
                    formats.Add(currentFile.Extension);
                    formatInstance.Add(1);
                    formatSize.Add(currentFile.Length);
                }
                else if(currentFile.Exists)
                {
                    formatInstance[formats.IndexOf(currentFile.Extension)] = formatInstance[formats.IndexOf(currentFile.Extension)] + 1;
                    formatSize[formats.IndexOf(currentFile.Extension)] = formatSize[formats.IndexOf(currentFile.Extension)] + currentFile.Length;
                }

                this.progressBar1.Invoke(updateProgressStep, new Object[] { i / 10 });

                try { dsoFile.Close(false); }
                catch (Exception ed) { };

            }

            //normalize common format variants
            #region normalize format variants
            object[] normalizeFormats = this.normalizeCommonFormatInstances(formats, formatInstance, formatSize);
            formats = (List<string>)normalizeFormats[0];
            formatInstance = (List<int>)normalizeFormats[1];
            formatSize = (List<float>)normalizeFormats[2];
            #endregion

            #region Compute Total Size
            //computing total root directory size
            for (int i = 0; i < formatSize.Count; i++)
                totalSize = totalSize + formatSize[i];

            string totalSizeString = "";

            if (totalSize > 134217728)
                totalSizeString = (totalSize / 134217728).ToString() + " GB";
            else if (totalSize > 1048576)
                totalSizeString = (totalSize / 1048576).ToString() + " MB";
            else if (totalSize > 1024)
                totalSizeString = (totalSize / 1024).ToString() + " B";
            #endregion

            #region Build Instance Graph
            //build graph
            ZedGraphControl zgc = new ZedGraphControl();
            GraphPane myPane = zgc.GraphPane;

            ZedGraphControl zgc2 = new ZedGraphControl();
            GraphPane myPane2 = zgc2.GraphPane;

            // Set the titles and axis labels
            myPane.Title.Text = "Formats Instance Graph";
            myPane.XAxis.Title.Text = "Format";
            myPane.YAxis.Title.Text = "Instances";

            myPane2.Title.Text = "Formats Size Graph";
            myPane2.XAxis.Title.Text = "Format";
            myPane2.YAxis.Title.Text = "Size in Byts";

            PointPairList list = new PointPairList();
            PointPairList list2 = new PointPairList();

            string[] labels = new string[formats.Count];

            for (int i = 0; i < formatInstance.Count; i++)
            {
                list.Add((double)i ,(double)formatInstance[i]);
                list2.Add((double)i, (double)formatSize[i]);
                labels[i] = formats[i];
            }

            // Generate a blue curve with circle symbols, and "My Curve 2" in the legend
            LineItem myCurve = myPane.AddCurve("Format By Instance", list, Color.Blue,SymbolType.Default);
            LineItem myCurve2 = myPane2.AddCurve("Format By Size", list2, Color.Blue, SymbolType.Default);

            // Fill the area under the curve with a white-red gradient at 45 degrees
            //myCurve.Line.Fill = new Fill(Color.White, Color.Red, 45F);

            // Make the symbols opaque by filling them with white
            myCurve.Symbol.Fill = new Fill(Color.White);
            myCurve2.Symbol.Fill = new Fill(Color.White);

            // Fill the axis background with a color gradient
            myPane.Chart.Fill = new Fill(Color.White, Color.LightGoldenrodYellow, 45F);
            myPane2.Chart.Fill = new Fill(Color.White, Color.LightGoldenrodYellow, 45F);

            // Fill the pane background with a color gradient
            myPane.Fill = new Fill(Color.White, Color.FromArgb(220, 220, 255), 45F);
            myPane2.Fill = new Fill(Color.White, Color.FromArgb(220, 220, 255), 45F);

            myPane.XAxis.Type = AxisType.Text;
            myPane2.XAxis.Type = AxisType.Text;

            myPane.XAxis.Scale.TextLabels = labels;
            myPane2.XAxis.Scale.TextLabels = labels;

            myPane.XAxis.Scale.FontSpec.Angle = 50;
            myPane2.XAxis.Scale.FontSpec.Angle = 50;

            // Calculate the Axis Scale Ranges
            zgc.AxisChange();
            zgc2.AxisChange();

            zgc.Height = 600;
            zgc.Width = 800;

            zgc2.Height = 600;
            zgc2.Width = 800;
            
            #endregion

            DirectoryInfo resultsDir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "\\results\\");
            if (!resultsDir.Exists)
                resultsDir.Create();

            String resultsFileName = DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + "_" + DateTime.Now.Hour + "'" + DateTime.Now.Minute + "'" + DateTime.Now.Second + "'" + DateTime.Now.Millisecond;

            FileStream filestream = new FileStream(resultsDir.FullName + "\\" + resultsFileName + ".html", FileMode.Create);
            StreamWriter resultsDoc = new StreamWriter(filestream);

            FileStream filestream2 = new FileStream(resultsDir.FullName + "\\" + resultsFileName + "_graphs.html", FileMode.Create);
            StreamWriter resultsDocAdd = new StreamWriter(filestream2);

            string instanceGraph = AppDomain.CurrentDomain.BaseDirectory + "results\\" + resultsFileName + "_ig.png";
            string sizeGraph = AppDomain.CurrentDomain.BaseDirectory + "results\\" + resultsFileName + "_sg.png";

            Image b = zgc.GetImage();
            b.Save(instanceGraph, System.Drawing.Imaging.ImageFormat.Png);

            Image c = zgc.GetImage();
            b.Save(sizeGraph, System.Drawing.Imaging.ImageFormat.Png);

            //zgc.SaveAs(instanceGraph);
            //zgc2.SaveAs(sizeGraph);

            DataTable dataSet = this.buildDataSet(formats, formatInstance, formatSize);
            dataSet.DefaultView.Sort = "["+dataSet.Columns[2].ColumnName + "] desc";
            dataSet = dataSet.DefaultView.ToTable();

            end = DateTime.Now;

            TimeSpan processTime = end.Subtract(start);

            d = new SetTextCallback(setStatusMsg);
            this.Invoke(d, new object[] { "Building results file" });

            object[] deepPathResults = this.deepestPath(allDirs);
            string deepestPath = (string)deepPathResults[0];
            int deepestPathDepth = (int)deepPathResults[1];

            resultsDoc.WriteLine("<b>Processing TimeSpan: </b>" + processTime.Days + ":" + processTime.Hours + ":" + processTime.Minutes + ":" + processTime.Seconds + ":" + processTime.Milliseconds);
            resultsDoc.WriteLine("<br><b>Number of Files: </b>" + allFiles.Count);
            resultsDoc.WriteLine("<br><b>Number of Potentially Dangerous Files: </b>" + this.dangerousFormatCount(formats,formatInstance));
            resultsDoc.WriteLine("<br><b>Number of Directories: </b>" + allDirs.Count);
            resultsDoc.WriteLine("<br><b>Deepest Directory of " + deepestPathDepth + " deep: </b>" + deepestPath);
            resultsDoc.WriteLine("<br><b>Average Folder Depth: </b>" + this.averageFolderDepth(allDirs));
            resultsDoc.WriteLine("<br><b>Number of Formats: </b>" + formats.Count);
            resultsDoc.WriteLine("<br><b>Total Size: </b>" + totalSizeString);
            if(fileAuthors.Count > 0)
                resultsDoc.WriteLine("<br><b>Biggest Contributor: </b>" + this.biggestContributor(fileAuthors,authorInstance));
            resultsDoc.WriteLine("<br><b>Oldest File: </b>\"" + oldestFileName + "\" <b>Created On:</b> " + oldestFileDate.ToLongDateString() + " " + oldestFileDate.ToLongTimeString());
            resultsDoc.WriteLine("<br><br><a href=\"file:///" + filestream2.Name + "\"><b>Graph Results</b></a><br><br>");

            resultsDoc.WriteLine("<table border='0'>");
            resultsDoc.WriteLine("<tr>");
            resultsDoc.WriteLine("<td>Formate</td>");
            resultsDoc.WriteLine("<td>Instances</td>");
            resultsDoc.WriteLine("<td>Size</td>");
            resultsDoc.WriteLine("</tr>");


            for (int i = 0; i < dataSet.Rows.Count; i++)
            {
                if ((float)dataSet.Rows[i].ItemArray[2] > 0)
                {
                    resultsDoc.WriteLine("<tr>");
                    resultsDoc.WriteLine("<td>" + dataSet.Rows[i].ItemArray[0] + "</td>");
                    resultsDoc.WriteLine("<td>" + dataSet.Rows[i].ItemArray[1] + "</td>");

                    float tmpSize = (float)dataSet.Rows[i].ItemArray[2];
                    string writeString = "";

                    if (tmpSize > 134217728)
                        writeString = (tmpSize / 134217728).ToString() + " GB";
                    else if (tmpSize > 1048576)
                        writeString = (tmpSize / 1048576).ToString() + " MB";
                    else if (tmpSize > 1024)
                        writeString = (tmpSize / 1024).ToString() + " KB";
                    else
                        writeString = tmpSize.ToString() + " B";

                    resultsDoc.WriteLine("<td>" + writeString + "</td>");
                    resultsDoc.WriteLine("</tr>");
                }
            }

            resultsDoc.WriteLine("</table>");
            resultsDoc.WriteLine("<br><b>Taxonomy</b><br><br>" + taxonomoy);

            this.currenResultsFile = filestream.Name;

            resultsDocAdd.WriteLine("<a href=\"file:///" + filestream.Name + "\">Back</a><br><br><b>By Format Instance Graph</b><br><img src=\"file:///" + instanceGraph + "\">");
            resultsDocAdd.WriteLine("<br><br><b>By Format Size Graph</b><br><img src=\"file:///" + sizeGraph + "\">");

            resultsDoc.Flush();
            resultsDocAdd.Flush();
            resultsDoc.Close();
            resultsDocAdd.Close();
            filestream.Close();
            filestream2.Close();

            GC.WaitForPendingFinalizers();
            GC.Collect();

            EnableResultsButton enableResults = new EnableResultsButton(EnableResultsBtn);
            this.button3.Invoke(enableResults);

        }

        public string BuildTaxonomy(List<string> allDirs, string root)
        {
            allDirs.Sort();
            string[] taxonomy = allDirs.ToArray();

            for (int i = 0; i < taxonomy.Length; i++)
            {
                taxonomy[i] = taxonomy[i].Replace(root, "");
                taxonomy[i] = taxonomy[i].Remove(0, 1);
            }

            //establish roots and depth
            for (int i = 0; i < taxonomy.Length; i++)
            {
                int depth = taxonomy[i].Split('\\').Length-1;

                if (taxonomy[i].Contains("\\"))
                    taxonomy[i] = taxonomy[i].Substring(taxonomy[i].LastIndexOf('\\')+1, taxonomy[i].Length - taxonomy[i].LastIndexOf('\\')-1);

                taxonomy[i] = taxonomy[i] + "*" + depth;
            }

            string htmlTaxonomy = root.Substring(root.LastIndexOf('\\')+1,root.Length-root.LastIndexOf('\\')-1 );

            for (int i = 0; i < taxonomy.Length; i++)
            {
                int upper = System.Convert.ToInt32(taxonomy[i].Substring(taxonomy[i].IndexOf('*')+1));
                htmlTaxonomy = htmlTaxonomy + "<br>";
                for( int j = 0; j <= upper; j++)
                    htmlTaxonomy = htmlTaxonomy + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";

                htmlTaxonomy = htmlTaxonomy + taxonomy[i].Substring(0,taxonomy[i].IndexOf('*'));
            }

            return htmlTaxonomy;
        }

        public int averageFolderDepth(List<string> dirs)
        {
            char[] evaluate = null;
            int average = 0;
            int runningtotal = 0;

            for (int i = 0; i < dirs.Count; i++)
            {
                evaluate = dirs[i].ToCharArray();

                for (int j = 0; j < evaluate.Length; j++)
                {
                    if (evaluate[j].Equals('\\'))
                        runningtotal = runningtotal + 1;
                }

                average = (average + runningtotal) / 2;
                runningtotal = 0;

            }

            return average;

        }

        public int dangerousFormatCount(List<string> formates, List<int> instances)
        {
            int count = 0;

            if (formates.Contains(".exe"))
                count = count + instances[formates.IndexOf(".exe")];
            if (formates.Contains(".dll"))
                count = count + instances[formates.IndexOf(".dll")];
            if (formates.Contains(".ocx"))
                count = count + instances[formates.IndexOf(".ocx")];
            if (formates.Contains(".dat"))
                count = count + instances[formates.IndexOf(".dat")];

            return count;
        }

        public object[] normalizeCommonFormatInstances(List<string> format, List<int> formatInstances, List<float> formatSize)
        {
            object[] results = new object[3];
            int index = 0;

            if( format.Contains(".TIF") && format.Contains(".tif") )
            {
                index = format.IndexOf(".tif");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".TIF")] = formatInstances[format.IndexOf(".TIF")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".TIF")] = formatSize[format.IndexOf(".TIF")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".TIF") && format.Contains(".Tif"))
            {
                index = format.IndexOf(".Tif");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".TIF")] = formatInstances[format.IndexOf(".TIF")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".TIF")] = formatSize[format.IndexOf(".TIF")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".TIF") && format.Contains(".tiff"))
            {
                index = format.IndexOf(".tiff");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".TIF")] = formatInstances[format.IndexOf(".TIF")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".TIF")] = formatSize[format.IndexOf(".TIF")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".TIF") && format.Contains(".TIFF"))
            {
                index = format.IndexOf(".TIFF");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".TIF")] = formatInstances[format.IndexOf(".TIF")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".TIF")] = formatSize[format.IndexOf(".TIF")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".bmp") && format.Contains(".BMP"))
            {
                index = format.IndexOf(".BMP");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".bmp")] = formatInstances[format.IndexOf(".bmp")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".bmp")] = formatSize[format.IndexOf(".bmp")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".cab") && format.Contains(".CAB"))
            {
                index = format.IndexOf(".CAB");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".cab")] = formatInstances[format.IndexOf(".cab")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".cab")] = formatSize[format.IndexOf(".cab")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".Cache") && format.Contains(".cache"))
            {
                index = format.IndexOf(".cache");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".Cache")] = formatInstances[format.IndexOf(".Cache")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".Cache")] = formatSize[format.IndexOf(".Cache")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".config") && format.Contains(".Config"))
            {
                index = format.IndexOf(".Config");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".config")] = formatInstances[format.IndexOf(".config")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".config")] = formatSize[format.IndexOf(".config")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".dll") && format.Contains(".DLL"))
            {
                index = format.IndexOf(".DLL");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".dll")] = formatInstances[format.IndexOf(".dll")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".dll")] = formatSize[format.IndexOf(".dll")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".doc") && format.Contains(".DOC"))
            {
                index = format.IndexOf(".DOC");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".doc")] = formatInstances[format.IndexOf(".doc")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".doc")] = formatSize[format.IndexOf(".doc")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".exe") && format.Contains(".EXE"))
            {
                index = format.IndexOf(".EXE");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".exe")] = formatInstances[format.IndexOf(".exe")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".exe")] = formatSize[format.IndexOf(".exe")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".exe") && format.Contains(".Exe"))
            {
                index = format.IndexOf(".Exe");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".exe")] = formatInstances[format.IndexOf(".exe")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".exe")] = formatSize[format.IndexOf(".exe")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".htm") && format.Contains(".HTM"))
            {
                index = format.IndexOf(".HTM");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".htm")] = formatInstances[format.IndexOf(".htm")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".htm")] = formatSize[format.IndexOf(".htm")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".htm") && format.Contains(".html"))
            {
                index = format.IndexOf(".html");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".htm")] = formatInstances[format.IndexOf(".htm")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".htm")] = formatSize[format.IndexOf(".htm")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".ico") && format.Contains(".ICO"))
            {
                index = format.IndexOf(".ICO");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".ico")] = formatInstances[format.IndexOf(".ico")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".ico")] = formatSize[format.IndexOf(".ico")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".inf") && format.Contains(".INF"))
            {
                index = format.IndexOf(".INF");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".inf")] = formatInstances[format.IndexOf(".inf")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".inf")] = formatSize[format.IndexOf(".inf")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".ini") && format.Contains(".INI"))
            {
                index = format.IndexOf(".INI");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".ini")] = formatInstances[format.IndexOf(".ini")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".ini")] = formatSize[format.IndexOf(".ini")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".ini") && format.Contains(".Ini"))
            {
                index = format.IndexOf(".Ini");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".ini")] = formatInstances[format.IndexOf(".ini")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".ini")] = formatSize[format.IndexOf(".ini")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".ini") && format.Contains(".ini2"))
            {
                index = format.IndexOf(".ini2");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".ini")] = formatInstances[format.IndexOf(".ini")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".ini")] = formatSize[format.IndexOf(".ini")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".jpg") && format.Contains(".jpeg"))
            {
                index = format.IndexOf(".jpeg");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".jpg")] = formatInstances[format.IndexOf(".jpg")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".jpg")] = formatSize[format.IndexOf(".jpg")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".jpg") && format.Contains(".JPG"))
            {
                index = format.IndexOf(".JPG");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".jpg")] = formatInstances[format.IndexOf(".jpg")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".jpg")] = formatSize[format.IndexOf(".jpg")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".log") && format.Contains(".LOG"))
            {
                index = format.IndexOf(".LOG");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".log")] = formatInstances[format.IndexOf(".log")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".log")] = formatSize[format.IndexOf(".log")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".log") && format.Contains(".Log"))
            {
                index = format.IndexOf(".Log");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".log")] = formatInstances[format.IndexOf(".log")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".log")] = formatSize[format.IndexOf(".log")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".msi") && format.Contains(".MSI"))
            {
                index = format.IndexOf(".MSI");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".msi")] = formatInstances[format.IndexOf(".msi")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".msi")] = formatSize[format.IndexOf(".msi")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".ocx") && format.Contains(".OCX"))
            {
                index = format.IndexOf(".OCX");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".ocx")] = formatInstances[format.IndexOf(".ocx")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".ocx")] = formatSize[format.IndexOf(".ocx")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".pdf") && format.Contains(".PDF"))
            {
                index = format.IndexOf(".PDF");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".pdf")] = formatInstances[format.IndexOf(".pdf")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".pdf")] = formatSize[format.IndexOf(".pdf")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".png") && format.Contains(".PNG"))
            {
                index = format.IndexOf(".PNG");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".png")] = formatInstances[format.IndexOf(".png")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".png")] = formatSize[format.IndexOf(".png")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".ppt") && format.Contains(".PPT"))
            {
                index = format.IndexOf(".PPT");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".ppt")] = formatInstances[format.IndexOf(".ppt")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".ppt")] = formatSize[format.IndexOf(".ppt")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".rar") && format.Contains(".RAR"))
            {
                index = format.IndexOf(".RAR");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".rar")] = formatInstances[format.IndexOf(".rar")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".rar")] = formatSize[format.IndexOf(".rar")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".resx") && format.Contains(".resX"))
            {
                index = format.IndexOf(".resX");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".resx")] = formatInstances[format.IndexOf(".resx")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".resx")] = formatSize[format.IndexOf(".resx")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".sav") && format.Contains(".SAV"))
            {
                index = format.IndexOf(".SAV");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".sav")] = formatInstances[format.IndexOf(".sav")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".sav")] = formatSize[format.IndexOf(".sav")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".scc") && format.Contains(".SCC"))
            {
                index = format.IndexOf(".SCC");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".scc")] = formatInstances[format.IndexOf(".scc")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".scc")] = formatSize[format.IndexOf(".scc")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".txt") && format.Contains(".TXT"))
            {
                index = format.IndexOf(".TXT");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".txt")] = formatInstances[format.IndexOf(".txt")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".txt")] = formatSize[format.IndexOf(".txt")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".wav") && format.Contains(".WAV"))
            {
                index = format.IndexOf(".WAV");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".wav")] = formatInstances[format.IndexOf(".wav")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".wav")] = formatSize[format.IndexOf(".wav")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".wmv") && format.Contains(".WMV"))
            {
                index = format.IndexOf(".WMV");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".wmv")] = formatInstances[format.IndexOf(".wmv")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".wmv")] = formatSize[format.IndexOf(".wmv")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".xls") && format.Contains(".XLS"))
            {
                index = format.IndexOf(".XLS");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".xls")] = formatInstances[format.IndexOf(".xls")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".xls")] = formatSize[format.IndexOf(".xls")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".xml") && format.Contains(".XML"))
            {
                index = format.IndexOf(".XML");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".xml")] = formatInstances[format.IndexOf(".xml")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".xml")] = formatSize[format.IndexOf(".xml")] + formatSize[index];
                formatSize.RemoveAt(index);

            }
            if (format.Contains(".zip") && format.Contains(".ZIP"))
            {
                index = format.IndexOf(".ZIP");

                format.RemoveAt(index);

                formatInstances[format.IndexOf(".zip")] = formatInstances[format.IndexOf(".zip")] + formatInstances[index];
                formatInstances.RemoveAt(index);

                formatSize[format.IndexOf(".zip")] = formatSize[format.IndexOf(".zip")] + formatSize[index];
                formatSize.RemoveAt(index);

            }

            results[0] = format;
            results[1] = formatInstances;
            results[2] = formatSize;

            return results;
        }

        public object[] deepestPath(List<string> dirs)
        {
            object[] results = new Object[2];
            string[] dirCopy = dirs.ToArray();
            char[] evaluate = null;
            int currentHigh = 0;
            int currentHighIndex = 0;

            for (int i = 0; i < dirCopy.Length; i++)
            {
                int size = dirCopy[i].Length;
                int localMax = 0;
                evaluate = dirCopy[i].ToCharArray();

                for (int j = 0; j < evaluate.Length; j++)
                {
                    if (evaluate[j].Equals('\\'))
                        localMax = localMax + 1;
                }

                if (localMax > currentHigh)
                {
                    currentHigh = localMax;
                    currentHighIndex = i;
                }
            }

            results[0] = dirs[currentHighIndex];
            results[1] = currentHigh;

            return results;

        }

        public string biggestContributor(List<string> authors, List<int> instances)
        {
            string biggestAuthor = authors[0];

            for (int i = 0; i < authors.Count - 1; i++)
                if (instances[i + 1] > instances[i])
                    biggestAuthor = authors[i + 1];

            return biggestAuthor;
        }
       
        public DataTable buildDataSet(List<string> listString,List<int> listInt,List<float> listFloat)
       {
           DataTable data = new DataTable();
           data.Columns.Add("Format", typeof(string));
           data.Columns.Add("Instances", typeof(int));
           data.Columns.Add("Size",typeof(float));

           for (int i = 0; i < listString.Count; i++)
           {
               data.Rows.Add(listString[i], listInt[i], listFloat[i]);
           }

           return data;

       }

        private void button3_Click(object sender, EventArgs e)
        {
            Process launchResults = new Process();
            Process.Start(this.currenResultsFile);
            this.button1.Enabled = true;
            this.button3.Enabled = false;
            this.pictureBox1.Image = this.imageList2.Images[1];
            this.pictureBox2.Image = this.imageList2.Images[1];
            this.setStatusMsg("Nothing");
            this.setFileMsg("Nothing");
        }

        private void button4_Click(object sender, EventArgs e)
        {

            AboutBox1 aboutSDA = new AboutBox1();
            aboutSDA.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FileInfo[] deleteThese = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "results\\").GetFiles();
            for (int i = 0; i < deleteThese.Length; i++)
                deleteThese[i].Delete();

            this.button5.Image = this.imageList3.Images[0];
        }
    }
}
