#region (C) Gary G. Jarkewicz 2008-2016
// 
// All rights are reserved. Reproduction or transmission in whole or in part, in
// any form or by any means, electronic, mechanical or otherwise, is prohibitie
// without the prior writtne permisison of the copyright owner.
//
// Filename: MainForm.cs
//
#endregion

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

/*
 * V1.3.0.0 6-Mar-2017  Add number of current titles on latest report after number of title in ().
 *                      Print message instead of lots of missing issues if more than 5 are missing for a title in the latest report.
 *                      Added a menu item under Process that will set the collection and investment values to the retail prices if used.
 * V1.4.0.0 15-Mar-2017 Add flag to only print titles with missing issues when printing the entire collection.
 * V1.5.0.0 1-Nov-2017  Changed text box for adding issues to up/down counter box for easier mouse use
 * V1.6.0.0 14-Jul-2018 Added Legacy and Update Year to database and code to quit printing missing issues for Legacy flagged comics due 
 *                          to large number jump for a while
 *                      Use Update Year to indicate when collection values were last updated for each title
 * V1.6.0.2 30-Mar-2022 Fix CSV export of title (names.csv) to remove duplicate first column and correct column headings.
 * 
 * */


namespace ComicCollector {
    public partial class MainForm : Form {
        public MainForm() {
            InitializeComponent();  // Standard Windows startup
            
            InitializeData();
        }

        enum reportPageLayout { Portrait, Landscape };
        enum reportTitlePerPage { Single, Multiple };

        /// <summary>
        /// To hold the currently selected title that is used by the report generator for a single title report.
        /// </summary>
        private static Title currentTitle = new Title();
        /// <summary>
        /// Holds whether the current report should generate the report across page boundaries or start each title on a separate page
        /// </summary>
        private static reportTitlePerPage currentReportTitlePerPage = reportTitlePerPage.Single;
        /// <summary>
        /// Length of the line for report generation (80 for portrait; 132 for landscape)
        /// </summary>
        private static int currentReportLineLength = 80;
        /// <summary>
        /// Whether to display the missing issues on a report
        /// </summary>
        private static bool printMissingIssues = true;
        /// <summary>
        /// Whether to only print titles with missing issues when printing the full collection.
        /// Will facilitate a printout more suitable for taking to a convention.
        /// </summary>
        private static bool onlyPrintTitlesWithMissingIssues = true;
        /// <summary>
        /// Maximum number of missing issues
        /// </summary>
        private static int maxMissingIssues = 2000;

        /// <summary>
        /// Routine to initialize data used by Comic Collector Database
        /// </summary>
        private void InitializeData() {
            // Look up directory to use for all data files (if set)
            string comicBookCollectionDirectory;

            // Check whether the environment variable exists.
            comicBookCollectionDirectory = Environment.GetEnvironmentVariable("ComicBookCollectionDirectory");
            // If necessary, create it.
            if (comicBookCollectionDirectory == null) {
                Environment.SetEnvironmentVariable("ComicBookCollectionDirectory", "c:\\comics\\");
                // Now retrieve it.
                comicBookCollectionDirectory = Environment.GetEnvironmentVariable("ComicBookCollectionDirectory");
            }
            Utilities.comicDatabaseFileName = comicBookCollectionDirectory + Utilities.comicDatabaseFileName;
            Utilities.outputFileName = comicBookCollectionDirectory + Utilities.outputFileName;
            Utilities.latestOutFileName = comicBookCollectionDirectory + Utilities.latestOutFileName;
            
            Utilities.verifyOutFileName = comicBookCollectionDirectory + Utilities.verifyOutFileName;
            Utilities.CSVFileForIssues = comicBookCollectionDirectory + Utilities.CSVFileForIssues;
            Utilities.CSVFileForNames = comicBookCollectionDirectory + Utilities.CSVFileForNames;
            // Set text boxes that show location names for output and latest listing files
            outputFileTextBox.Text = Utilities.outputFileName;
            latestFileTextBox.Text = Utilities.latestOutFileName;
        }


        /// <summary>
        /// Routine to handle Row Enter event for the Title data grid.
        /// Fills in the Issue data grid with all the information for this title along with the
        /// missing issue area.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TitleGridView_RowEnter(object sender, DataGridViewCellEventArgs e) {
            bool[] missingIssues = new bool[maxMissingIssues];
            string missingIssueText = "";
            int j, numberMissing;
            decimal totalCollectionValue = 0.0M;
            decimal totalInvestmentValue = 0.0M;
            decimal totalRetailPrice = 0.0M;
            int index = 0, numberOfIssues = 0;
            Title localTitle = new Title();
            Issue localIssue = new Issue();

            IssueGridView.Rows.Clear();
            comicStatusLabel.Text = "Reading Issues";
            for (j = 0; j < maxMissingIssues; j++) {
                missingIssues[j] = true;
            }
            if (TitleGridView.CurrentRow != null) {
                // causes an exception but required to fill in issue information
                try {
                    index = Convert.ToInt16(TitleGridView.SelectedRows[0].Cells[0].Value);
                } catch (Exception ex) {
                    Console.WriteLine("Error: " + ex.Message);
                }
            } else {
                return;
            }
            // Fill in currentIssues array from database
            numberOfIssues = Utilities.fillCurrentIssues(index);

            for (int i = 0; i < numberOfIssues; i++) {
                localIssue = Utilities.currentIssues[i];
                IssueGridView.Rows.Add(index, localIssue.issueNumber, localIssue.condition, localIssue.retailPrice,
                localIssue.investmentValue, localIssue.collectionValue);
                // If the issue number is a clean (no decimal point) number, then indicate we have it. This will ignore x.1, x.2 issues
                missingIssues[(int)Math.Floor(localIssue.issueNumber)] = false;
                if (localIssue.collectionValue == 0.0M) {
                    totalCollectionValue += localIssue.retailPrice;
                } else {
                    totalCollectionValue += localIssue.collectionValue;
                }
                if (localIssue.investmentValue == 0.0M) {
                    totalInvestmentValue += localIssue.retailPrice;
                } else {
                    totalInvestmentValue += localIssue.investmentValue;
                }
                totalRetailPrice += Convert.ToDecimal(localIssue.retailPrice);
            }

            if (numberOfIssues > 1) {
                IssueGridView.Sort(IssueGridView.Columns[1], ListSortDirection.Ascending);  // Sort the Issue grid by the Issue number for better appearance
            }

            try {
                CurrentTitleText.Text = Convert.ToString(TitleGridView.SelectedRows[0].Cells["Title"].Value);
                currentTitle.title = Convert.ToString(TitleGridView.SelectedRows[0].Cells["Title"].Value);
                currentTitle.type = Convert.ToInt16(TitleGridView.SelectedRows[0].Cells["Type"].Value);
                currentTitle.publisher = Convert.ToString(TitleGridView.SelectedRows[0].Cells["Publisher"].Value);
                currentTitle.index = Convert.ToInt16(TitleGridView.SelectedRows[0].Cells[0].Value);
                currentTitle.lastIssue = Convert.ToInt16(TitleGridView.SelectedRows[0].Cells["LastIssue"].Value);
                currentTitle.numberOfIssues = Convert.ToInt16(TitleGridView.SelectedRows[0].Cells["NumberOfIssues"].Value);
                currentTitle.legacy = Convert.ToBoolean(TitleGridView.SelectedRows[0].Cells["Legacy"].Value);
                currentTitle.updateyear = Convert.ToInt16(TitleGridView.SelectedRows[0].Cells["UpdateYear"]);
            } catch (Exception ex) {
                Console.WriteLine("Error accessing SelectedRows[0]" + ex.Message);
            }

            // TODO: Check this for what happens with an x.1, x.2 issue in the mix - seems to work OK.
            for (j = 1, numberMissing = 0; j < currentTitle.lastIssue; j++) {
                if (missingIssues[j]) {
                    missingIssueText += Convert.ToString(j) + ", ";
                    numberMissing++;
                }
            }
            MissingIssuesTotalText.Text = numberMissing.ToString();
            RetailPriceText.Text = totalRetailPrice.ToString();
            InvestmentValueText.Text = totalInvestmentValue.ToString();
            MissingIssuesText.Text = missingIssueText;
            CollectionValueText.Text = totalCollectionValue.ToString();

            comicStatusLabel.Text = "Done";
            this.Cursor = Cursors.Arrow;
        }


        /// <summary>
        /// Event handler for the Exit menu item.
        /// Closes the application.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exitToolStripMenuItem_Click(object sender, EventArgs e) {
            this.titlesTableAdapter.Update(this.comicDataBaseDataSet1.Titles);
            this.Close();
        }

        /// <summary>
        /// Update database from dataset
        /// </summary>
        private void exitComicCollector(object sender, EventArgs e) {
            this.titlesTableAdapter.Update(this.comicDataBaseDataSet1.Titles);
        }

        /// <summary>
        /// Event handler for the About menu item.
        /// Displays the version information.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e) {
            MessageBox.Show("Comic Collector V1.6.0.2\r\n22-Mar-2022\r\nCopyright 2008-2022\r\nGary Jarkewicz");
        }

        // When a character is typed into a row (or the grid is built up)
        private void IssueGridView_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e) {
            //MessageBox.Show("Row added");
        }

        // When the row is selected
        private void IssueGridView_RowEnter(object sender, DataGridViewCellEventArgs e) {
            //MessageBox.Show("Issue row entered");
        }

        private void printSummaryOfCollection(StreamWriter SW) {
            decimal totalRetailPrice = 0;
            int totalNumberOfIssues = 0;
            int numberCurrentTitles = 0;
            decimal totalInvestmentValue = 0.0M;
            decimal totalCollectionValue = 0.0M;
            int numberOfTitles = TitleGridView.RowCount;
            string outputLine;

            generateSummary(ref numberOfTitles, ref numberCurrentTitles, ref totalNumberOfIssues, ref totalRetailPrice,
                ref totalInvestmentValue, ref totalCollectionValue);

            outputLine = "\tTotal Number of Titles: " + numberOfTitles.ToString() + " (" + numberCurrentTitles.ToString() + " current)" + "\tTotal Number of Issues: " + totalNumberOfIssues.ToString();
            SW.WriteLine(outputLine);
            outputLine = String.Format("\tTotal Retail Price:\t{0:C}", totalRetailPrice) + String.Format("\t\tTotal Investment Value:\t{0:C}", totalInvestmentValue);
            SW.WriteLine(outputLine);
            outputLine = String.Format("\tTotal Collection Value:\t{0:C}", totalCollectionValue);
            SW.WriteLine(outputLine);
            SW.WriteLine();
            SW.WriteLine();
        }

        /// <summary>
        /// Routine to print the summary of the latest issues of the selected type.
        /// </summary>
        /// <param name="type"></param>
        private void printLatestIssueSummary(int type) {
            StreamWriter SW;
            SW = File.CreateText(Utilities.latestOutFileName);
            int numberOfIssues = 0;
            bool[] missingIssues = new bool[maxMissingIssues];
            string missingIssueText = "";
            int j;
            decimal retailPriceOfLastIssue;
            Title localTitle = new Title();
            Issue localIssue = new Issue();
            int numberOfMissingIssues = 0;

            this.Cursor = Cursors.WaitCursor;

            // Sort the TitleGridView by the Title name before printing
            TitleGridView.Sort(TitleGridView.Columns[1], ListSortDirection.Ascending);  // Sort the Title grid by the Title Name (Column 1)

            int numberOfTitles = TitleGridView.RowCount;
            comicStatusLabel.Text = "Printing Latest Issue Report";
            string outputLine = "Latest Issues Report Generated on " + System.DateTime.Now.ToString();

            SW.WriteLine(outputLine.PadLeft((outputLine.Length / 2) + (currentReportLineLength / 2)));
            SW.WriteLine();
            SW.WriteLine();

            printSummaryOfCollection(SW);

            for (j = 0; j < maxMissingIssues; j++) {
                missingIssues[j] = true;
            }
            
            for (int title = 0; title < numberOfTitles; title++) {
                localTitle = Utilities.ReadTitle(TitleGridView.Rows, title);
                if (((localTitle.type & type) != 0) ||
                    (type == 4)) {
                    // Get issue data for this title to get last issue price
                    numberOfIssues = Utilities.fillCurrentIssues(localTitle.index);
                    retailPriceOfLastIssue = Utilities.currentIssues[numberOfIssues - 1].retailPrice;
                    // Use ToString("F2") for formatting Retail Price from last issue from this title
                    SW.WriteLine("{0,5}  {1,-45}{2,5}{3,5}{4,5} ({5,5})", Convert.ToString(localTitle.index), localTitle.title,
                        localTitle.publisher, localTitle.type.ToString(), Convert.ToString(localTitle.lastIssue), retailPriceOfLastIssue.ToString("F2"));


                    if (localTitle.lastIssue > localTitle.numberOfIssues) {
                        numberOfMissingIssues = 0;
                        missingIssueText = "";
                        for (int i = 0; i < numberOfIssues; i++) {
                            localIssue = Utilities.currentIssues[i];
                            // If the issue number is a clean (no decimal point) number, then indicate we have it. This will ignore x.1, x.2 issues
                            missingIssues[(int)Math.Floor(localIssue.issueNumber)] = false;
                        }

                        for (j = 1; j < localTitle.lastIssue; j++) {
                            if (missingIssues[j]) {
                                missingIssueText += Convert.ToString(j) + ", ";
                                numberOfMissingIssues++;
                            }
                        }
                        if (numberOfMissingIssues <= 5) {
                            SW.WriteLine("            Missing: " + missingIssueText);
                        } else {
                            if (!localTitle.legacy) {   // If a legacy run, don't print this due to large issue number jump
                                SW.WriteLine("            More than 5 Missing Issues");
                            }
                            
                        }
                    }
                }
            }







            SW.Close();
            this.Cursor = Cursors.Arrow;
            comicStatusLabel.Text = "Done";
        }

        /// <summary>
        /// Connected to Selected (or Current) Title Report menu item.
        /// Will use currentTitle which is set when a row from the Title data grid is selected to pass to the
        /// outputTitle() routine. This will, in turn, output to the output file a report on this title.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void currentTitleReportToolStripMenuItem_Click(object sender, EventArgs e) {
            StreamWriter SW;
            SW = File.CreateText(Utilities.outputFileName);
            outputTitle(currentTitle, SW, true, false);
            SW.Close();
        }

        /// <summary>
        /// Connected to both Title File Report and Title File Export menu item. Depending on which item sent this event, the output file 
        /// is adjusted for proper alignment for text (Report) or importing into Excel (Export)
        /// </summary>
        /// <param name="sender">Menu Item that is connected to this event.</param>
        /// <param name="e"></param>
        private void titleFileReportToolStripMenuItem_Click(object sender, EventArgs e) {
            StreamWriter SW;
            SW = File.CreateText(Utilities.outputFileName);
            Title localTitle = new Title();
            this.Cursor = Cursors.WaitCursor;
            int numberOfTitles = TitleGridView.RowCount;
            comicStatusLabel.Text = "Printing Title File Report";
            string outputLine = "Title File Report Generated on " + System.DateTime.Now.ToString();

            SW.WriteLine(outputLine.PadLeft((outputLine.Length / 2) + (currentReportLineLength / 2)));
            SW.WriteLine();
            SW.WriteLine();
            outputLine = "Record\t".PadLeft(10) +
                "Title".PadLeft(("Title".Length / 2) + (Utilities.titleSectionLength / 2)) +
                " ".PadLeft(Utilities.titleSectionLength / 2) +
                "\tPublisher\t".PadLeft(10) + "Type\t" + "Last Issue\t".PadLeft(10) +
                "Number\t".PadLeft(10);
            if (printMissingIssues) {
                outputLine += "\tMissing";
            }
            SW.WriteLine(outputLine);
            for (int title = 0; title < numberOfTitles; title++) {
                localTitle = Utilities.ReadTitle(TitleGridView.Rows, title);

                if (sender == titleFileExportToolStripMenuItem) {
                    outputLine = Convert.ToString(title).PadLeft(10) +
                        "\t" + localTitle.title + "\t" + localTitle.publisher.PadLeft(10) +
                        "\t" + Convert.ToString(localTitle.type) +
                        "\t" + Convert.ToString(localTitle.lastIssue).PadLeft(10) +
                        "\t" + Convert.ToString(localTitle.numberOfIssues).PadLeft(10);
                } else {
                    outputLine = Convert.ToString(title).PadLeft(10) +
                        "\t" + localTitle.title + " " + localTitle.publisher.PadLeft(10) +
                        "\t\t" + Convert.ToString(localTitle.type) +
                        "\t" + Convert.ToString(localTitle.lastIssue).PadLeft(10) +
                        "\t" + Convert.ToString(localTitle.numberOfIssues).PadLeft(10);
                }
                if (printMissingIssues) {
                    outputLine += "\t" + Convert.ToString(localTitle.lastIssue - localTitle.numberOfIssues);
                }
                SW.WriteLine(outputLine);
            }
            SW.Close();
            this.Cursor = Cursors.Arrow;
            comicStatusLabel.Text = "Done";
        }

        /// <summary>
        /// Scan thru issues database and compute the totals requested.
        /// </summary>
        /// <param name="numberOfTitles">Returned number of title - computed from size of table.</param>
        /// <param name="numberCurrentTitles">Returned number of current titles.</param>
        /// <param name="totalNumberOfIssues">Returned number of issues.</param>
        /// <param name="totalRetailPrice">Returned total retail price of all issues.</param>
        /// <param name="totalInvestmentValue">Returned total investment value of all issues.</param>
        /// <param name="totalCollectionValue">Returned total collection value of all issues.</param>
        private void generateSummary(ref int numberOfTitles, ref int numberCurrentTitles, ref int totalNumberOfIssues,
            ref decimal totalRetailPrice, ref decimal totalInvestmentValue, ref decimal totalCollectionValue) {
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Utilities.comicDatabaseFileName;
            Issue localIssue = new Issue();
            totalRetailPrice = 0.0M;
            totalNumberOfIssues = 0;
            totalInvestmentValue = 0.0M;
            totalCollectionValue = 0.0M;

            numberOfTitles = TitleGridView.RowCount;
            for (int i = 0; i < numberOfTitles - 1; i++) {
                if ((Convert.ToInt16(TitleGridView.Rows[i].Cells["Type"].Value) & 1) != 0) {
                    numberCurrentTitles++;
                }
            }

            string strAccessSelect = "SELECT * FROM Issues";
            // Create the dataset and add the Categories table to it:
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try {
                myAccessConn = new OleDbConnection(strAccessConn);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }
            try {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(myDataSet);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            } finally {
                myAccessConn.Close();
            }

            DataRowCollection dra = myDataSet.Tables["Table"].Rows;
            foreach (DataRow dr in dra) {
                localIssue = Utilities.ReadIssue(dr);
                totalRetailPrice += localIssue.retailPrice;
                totalNumberOfIssues++;
                if (localIssue.investmentValue == 0.0M) {
                    totalInvestmentValue += localIssue.retailPrice;
                } else {
                    totalInvestmentValue += localIssue.investmentValue;
                }
                if (localIssue.collectionValue == 0.0M) {
                    totalCollectionValue += localIssue.retailPrice;
                } else {
                    totalCollectionValue += localIssue.collectionValue;
                }
            }
        }

        /// <summary>
        /// Generate a summary report of the entire collection to the outputFileName
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void summaryReportToolStripMenuItem_Click(object sender, EventArgs e) {
            StreamWriter SW;
            SW = File.CreateText(Utilities.outputFileName);
            decimal totalRetailPrice = 0;
            int totalNumberOfIssues = 0;
            int numberCurrentTitles = 0;
            decimal totalInvestmentValue = 0.0M;
            decimal totalCollectionValue = 0.0M;
            this.Cursor = Cursors.WaitCursor;
            int numberOfTitles = 0;
            comicStatusLabel.Text = "Printing Summary Report";
            string outputLine = "Summary Report Generated on " + System.DateTime.Now.ToString();

            SW.WriteLine(outputLine.PadLeft((outputLine.Length / 2) + (currentReportLineLength / 2)));
            SW.WriteLine();
            SW.WriteLine();

            generateSummary(ref numberOfTitles, ref numberCurrentTitles, ref totalNumberOfIssues, ref totalRetailPrice,
                ref totalInvestmentValue, ref totalCollectionValue);

            outputLine = "\tTotal Number of Titles: " + numberOfTitles.ToString();
            SW.WriteLine(outputLine);
            outputLine = "\tTotal Number of Issues: " + totalNumberOfIssues.ToString();
            SW.WriteLine(outputLine);
            outputLine = String.Format("\tTotal Retail Price:\t{0:C}", totalRetailPrice);
            SW.WriteLine(outputLine);
            outputLine = String.Format("\tTotal Investment Value:\t{0:C}", totalInvestmentValue);
            SW.WriteLine(outputLine);
            outputLine = String.Format("\tTotal Collection Value:\t{0:C}", totalCollectionValue);
            SW.WriteLine(outputLine);
            SW.Close();
            this.Cursor = Cursors.Arrow;
            comicStatusLabel.Text = "Done";
        }

        /// <summary>
        /// Set of routines connected to menu items to set various internal flags for report behaviors
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void portraitToolStripMenuItem_Click(object sender, EventArgs e) {
            currentReportLineLength = Utilities.portraitLineLength;
        }

        private void landscapeToolStripMenuItem_Click(object sender, EventArgs e) {
            currentReportLineLength = Utilities.landscapeLineLength;
        }

        private void singleToolStripMenuItem_Click(object sender, EventArgs e) {
            currentReportTitlePerPage = reportTitlePerPage.Single;
        }

        private void multipleToolStripMenuItem_Click(object sender, EventArgs e) {
            currentReportTitlePerPage = reportTitlePerPage.Multiple;
        }

        private void currentToolStripMenuItem_Click(object sender, EventArgs e) {
            printLatestIssueSummary(1);
        }

        private void completeToolStripMenuItem_Click(object sender, EventArgs e) {
            printLatestIssueSummary(2);
        }

        private void currentAndCompleToolStripMenuItem_Click(object sender, EventArgs e) {
            printLatestIssueSummary(3);
        }

        private void allToolStripMenuItem_Click(object sender, EventArgs e) {
            printLatestIssueSummary(4);
        }

        private void yesToolStripMenuItem_Click(object sender, EventArgs e) {
            printMissingIssues = true;
        }

        private void noToolStripMenuItem_Click(object sender, EventArgs e) {
            printMissingIssues = false;
        }

        private void yesToolStripMenuItem1_Click(object sender, EventArgs e) {
            onlyPrintTitlesWithMissingIssues = true;
        }

        private void noToolStripMenuItem1_Click(object sender, EventArgs e) {
            onlyPrintTitlesWithMissingIssues = false;
        }

        /// <summary>
        /// Output a report on the currently selected title. List the title and its information along with all 
        /// the issues for this title to the current output file.
        /// </summary>
        /// <param name="localTitle">Currently selected title to report.</param>
        private void outputTitle(Title localTitle, StreamWriter SW, bool printDateOfReport, bool onlyIfMissingIssues) {
            bool[] missingIssues = new bool[maxMissingIssues];
            decimal totalCollectionValue = 0.0M;
            decimal totalInvestmentValue = 0.0M;
            decimal totalRetailPrice = 0.0M;
            int j, numberMissing = 0, numberOfIssues = 0;
            int index = 0;
            Issue localIssue = new Issue();
            decimal localRetailPrice;
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Utilities.comicDatabaseFileName;
            string strAccess = "";

            for (j = 0; j < maxMissingIssues; j++) {
                missingIssues[j] = true;
            }

            this.Cursor = Cursors.WaitCursor;
            comicStatusLabel.Text = "Printing Title Report: " + localTitle.title;

            // Report title line
            string outputLine = "Title: " + localTitle.title;
            if (printDateOfReport) {
                outputLine = outputLine + " Report Generated on " + System.DateTime.Now.ToString();
            }

            // If there are no missing issues and not to print, then skip
            if ((localTitle.numberOfIssues == localTitle.lastIssue) && (onlyIfMissingIssues)) {
                return;
            }


            SW.WriteLine(outputLine.PadLeft((outputLine.Length / 2) + (currentReportLineLength / 2)));
            SW.WriteLine();
            SW.WriteLine("\t\t\t\t()= Investment; [] = Collection; <> = Condition");
            SW.WriteLine();

            outputLine = "\t" + localTitle.title + "\t\t(PC: " + localTitle.publisher + ")" +
                "\t(Type: " + localTitle.type.ToString() + ")";
            SW.WriteLine(outputLine);
            SW.WriteLine();

            // Read all the issues for this title and generate the report information for them
            
            index = localTitle.index;

            strAccess = "SELECT * FROM Issues where index = " + index.ToString() + " ORDER BY Issue";

            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try {
                myAccessConn = new OleDbConnection(strAccessConn);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }
            try {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccess, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(myDataSet);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            } finally {
                myAccessConn.Close();
            }

            DataRowCollection dra = myDataSet.Tables["Table"].Rows;
            localIssue = Utilities.ReadIssue(dra[0]);
            localRetailPrice = localIssue.retailPrice;
            outputLine = new String(' ', 16);   // equivalent to 2 tabs
            foreach (DataRow dr in dra) {
                localIssue = Utilities.ReadIssue(dr);
                missingIssues[(int)Math.Floor(localIssue.issueNumber)] = false;
                if (localIssue.collectionValue == 0.0M) {
                    totalCollectionValue += localIssue.retailPrice;
                } else {
                    totalCollectionValue += localIssue.collectionValue;
                }
                if (localIssue.investmentValue == 0.0M) {
                    totalInvestmentValue += localIssue.retailPrice;
                } else {
                    totalInvestmentValue += localIssue.investmentValue;
                }
                totalRetailPrice += Convert.ToDecimal(localIssue.retailPrice);
                numberOfIssues++;

                outputLine = outputLine + localIssue.printout + ", ";
                if ((outputLine.Length > (currentReportLineLength - 35)) | (localRetailPrice != localIssue.retailPrice)) {
                    outputLineToFile(outputLine, localRetailPrice, SW);
                    outputLine = new String(' ', 16);   // equivalent to 2 tabs
                    if (localRetailPrice != localIssue.retailPrice) {
                        localRetailPrice = localIssue.retailPrice;
                    }
                }
            }
            outputLineToFile(outputLine, localRetailPrice, SW);
            SW.WriteLine();
            SW.WriteLine();
            string filler = new String(' ', currentReportLineLength - 50);
            outputLine = filler + "Number of Issues " + numberOfIssues;
            SW.WriteLine(outputLine);
            outputLine = filler + "Retail Price     " + Convert.ToString(totalRetailPrice);
            SW.WriteLine(outputLine);
            outputLine = filler + "Investment Value " + Convert.ToString(totalInvestmentValue);
            SW.WriteLine(outputLine);
            outputLine = filler + "Collection Value " + Convert.ToString(totalCollectionValue);
            SW.WriteLine(outputLine);

            if (printMissingIssues & (numberOfIssues != localTitle.lastIssue)) {
                outputLine = filler + "Number of Missing Issues " + Convert.ToString(localTitle.lastIssue - numberOfIssues);
                SW.WriteLine(outputLine);
                SW.WriteLine();
                outputLine = new String(' ', 10);
                for (j = 1; j < localTitle.lastIssue; j++) {
                    if (missingIssues[j] == true) {
                        numberMissing++;
                        outputLine = outputLine + Convert.ToString(j) + ", ";
                        if (outputLine.Length > (currentReportLineLength)) {
                            SW.WriteLine(outputLine);
                            SW.WriteLine();
                            outputLine = new String(' ', 10);
                        }
                    }
                }
                SW.WriteLine(outputLine);
                SW.WriteLine();
            }
            // Check for putting a form feed after the page is output
            if (currentReportTitlePerPage == reportTitlePerPage.Single) {
                SW.Write("\f");
            } else {
                outputLine = new String('_', 70);
                SW.WriteLine(outputLine);
                SW.WriteLine();
            }

            this.Cursor = Cursors.Arrow;
            comicStatusLabel.Text = "Done";
        }

        /// <summary>
        /// Put the supplied line (ol) to the output file after appending the retail price information with
        /// proper spacing.
        /// </summary>
        /// <param name="ol">Current line to output</param>
        /// <param name="rp">Retail price</param>
        /// <param name="SW">StreamWriter to use for output</param>
        private void outputLineToFile(string ol, decimal rp, StreamWriter SW) {
            string filler;
            int fillerLength = (currentReportLineLength - 10) - ol.Length;
            if (fillerLength >= 0) {
                filler = new String(' ', fillerLength);
            } else {
                filler = "";
            }
            ol = ol + filler + "$";
            ol = ol + Convert.ToString(rp);
            SW.WriteLine(ol);
        }


        /// <summary>
        /// Connected to the Full Collection Report Menu Item.
        /// Dump the entire collection to the output file indicated by the outputFileName 'global'.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void fullCollectionMenuItem_Click(object sender, EventArgs e) {
            StreamWriter SW;
            Title localTitle = new Title();
            int numberOfTitles = TitleGridView.RowCount;
            this.Cursor = Cursors.WaitCursor;
            SW = File.CreateText(Utilities.outputFileName);
            comicStatusLabel.Text = "Printing Titles";
            string outputLine = "Report Generated on " + System.DateTime.Now.ToString() + ":";
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Utilities.comicDatabaseFileName;

            // Sort the TitleGridView by the Title name before printing
            TitleGridView.Sort(TitleGridView.Columns[1], ListSortDirection.Ascending);  // Sort the Title grid by the Title Name (Column 1)

            SW.WriteLine(outputLine);
            SW.WriteLine();

            printSummaryOfCollection(SW);
            SW.Write("\f");

            string strAccessSelect = "SELECT * FROM Titles ORDER BY Title";
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try {
                myAccessConn = new OleDbConnection(strAccessConn);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }
            try {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(myDataSet);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            } finally {
                myAccessConn.Close();
            }

            DataRowCollection dra = myDataSet.Tables["Table"].Rows;
            foreach (DataRow dr in dra) {
                try {
                    localTitle = Utilities.ReadTitle(dr);
                    Console.WriteLine("Processing: " + localTitle.title);
                    outputTitle(localTitle, SW, false, onlyPrintTitlesWithMissingIssues);
                    Application.DoEvents();
                } catch (Exception ex) {
                    comicStatusLabel.Text = "Error: " + ex.Message;
                    break;
                }
            }

            comicStatusLabel.Text = "Done";
            SW.Close();

            this.Cursor = Cursors.Arrow;
        }
 
        /// <summary>
        /// Attached to the text changing in the output file name text box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void outputFileTextBox_TextChanged(object sender, EventArgs e) {
            Utilities.outputFileName = outputFileTextBox.Text;
        }

        /// <summary>
        /// Attached to the text changing in the latest file name text box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void latestFileTextBox_TextChanged(object sender, EventArgs e) {
            Utilities.latestOutFileName = latestFileTextBox.Text;
        }

        /// <summary>
        /// Tests the PCPrint class by printing out the summary of the collection to the default printer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void testPrintClassToolStripMenuItem_Click(object sender, EventArgs e) {
            // Create an instance of our printer class
            PCPrint printer = new PCPrint();
            // Set the font we want to use
            printer.PrinterFont = new Font("Verdana", 10);

            string textToPrint = "";
            decimal totalRetailPrice = 0;
            int totalNumberOfIssues = 0;
            int numberCurrentTitles = 0;
            decimal totalInvestmentValue = 0.0M;
            decimal totalCollectionValue = 0.0M;
            this.Cursor = Cursors.WaitCursor;
            int numberOfTitles = 0;
            comicStatusLabel.Text = "Printing Summary Report";
            string outputLine = "Summary Report Generated on " + System.DateTime.Now.ToString();

            textToPrint = outputLine.PadLeft((outputLine.Length / 2) + (currentReportLineLength / 2)) + "\r\n\r\n\r\n";
            generateSummary(ref numberOfTitles, ref numberCurrentTitles, ref totalNumberOfIssues, ref totalRetailPrice,
                ref totalInvestmentValue, ref totalCollectionValue);

            textToPrint += "\tTotal Number of Titles:   " + String.Format("{0:N0}", numberOfTitles) + "\r\n";
            textToPrint += "\t" + String.Format("Total Number of Issues: {0:N0}", totalNumberOfIssues) + "\r\n";
            textToPrint += String.Format("\tTotal Retail Price:\t  {0:C}", totalRetailPrice) + "\r\n";
            textToPrint += String.Format("\tTotal Investment Value:\t  {0:C}", totalInvestmentValue) + "\r\n";
            textToPrint += String.Format("\tTotal Collection Value:\t  {0:C}", totalCollectionValue) + "\r\n";
            printer.TextToPrint = textToPrint;
            printer.Print();

            this.Cursor = Cursors.Arrow;
            comicStatusLabel.Text = "Done";
        }

        /// <summary>
        /// Exports the data files to a CSV format for import into an SQL database. 
        /// Somewhat redundant since we're using an SQL database, but just in case we want to use Excel for any processing.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exportToCSVToolStripMenuItem_Click(object sender, EventArgs e) {
            StreamWriter SWNames, SWIssues;
            Title localTitle = new Title();
            int numberOfTitles = TitleGridView.RowCount;
            Issue localIssue = new Issue();
            string outputLine = "";
            this.Cursor = Cursors.WaitCursor;
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Utilities.comicDatabaseFileName;

            SWNames = File.CreateText(Utilities.CSVFileForNames);
            SWIssues = File.CreateText(Utilities.CSVFileForIssues);
            comicStatusLabel.Text = "Exporting Titles and Issues";
            outputLine = localIssue.CSVHeaderOut;
            SWIssues.WriteLine(outputLine);
            outputLine = localTitle.CSVHeaderOut;
            SWNames.WriteLine(outputLine);

            // First process the titles to that file
            string strAccessSelect = "SELECT * FROM Titles ORDER BY Title";
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try {
                myAccessConn = new OleDbConnection(strAccessConn);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }
            try {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(myDataSet);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            } finally {
                myAccessConn.Close();
            }

            DataRowCollection dra = myDataSet.Tables["Table"].Rows;
            foreach (DataRow dr in dra) {
                localTitle = Utilities.ReadTitle(dr);
                outputLine = localTitle.CSVOut;
                SWNames.WriteLine(outputLine);
            }

            // Then the issues to that file.
            strAccessSelect = "SELECT * FROM Issues";
            // Create the dataset and add the Categories table to it:
            myDataSet = new DataSet();
            myAccessConn = null;
            try {
                myAccessConn = new OleDbConnection(strAccessConn);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }
            try {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(myDataSet);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            } finally {
                myAccessConn.Close();
            }

            dra = myDataSet.Tables["Table"].Rows;
            foreach (DataRow dr in dra) {
                localIssue = Utilities.ReadIssue(dr);
                outputLine = localIssue.CSVOut();
                SWIssues.WriteLine(outputLine);
            }

            comicStatusLabel.Text = "Done";
            SWNames.Close();
            SWIssues.Close();

            this.Cursor = Cursors.Arrow;
        }

        /// <summary>
        /// Load of the MainForm for the ComicCollectorDatabase - Loads the Titles from the Access Database which is Data Bound to the Titles Data Grid
        /// Due to the data binding any updates to the form on the screen will be set in the database on exit of the program due to code added to 
        /// update the database on program exit or menu selection to save the names.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e) {
            // This line of code loads data into the 'comicDataBaseDataSet.Titles' table. You can move, or remove it, as needed.
            this.titlesTableAdapter.Fill(this.comicDataBaseDataSet1.Titles); // Fill the Datagrid on the screen from the Titles database file
            TitleGridView.Sort(TitleGridView.Columns[0], ListSortDirection.Ascending);  // Sort the Title grid by the Index to facilitate adding titles (Column 0)
            Console.WriteLine("MainForm_Load");
            // Set the missing index values to -1 for later lookup and use
            for (int i = 0; i < Utilities.maxTitleLimit; i++) {
                Utilities.missingTitleIndex[i] = -1;
            }
            // Set valid title index values - removed with latest update to database (14-Jul)
            for (int i = 0; i < TitleGridView.NewRowIndex; i++) {
                Utilities.missingTitleIndex[Convert.ToInt16(TitleGridView.Rows[i].Cells[0].Value)] = i + 1;
            }
            // Tht titles will be color coded for the background of titles based on complete or not in the event that handles
            // the format changed event.
        }

        /// <summary>
        /// Update Issue Database Record from Datagrid information on the screen.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void IssueAddOrUpdateButton_Click(object sender, EventArgs e) {
            // Provide code to do a Delete based on the Index, then an Insert from the datagrid using SQL commands to the Issues Table
            // Delete the issue first
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Utilities.comicDatabaseFileName;
            string strAccess = "";
            int index = 0;
            decimal retailPrice, investmentValue, collectionValue;
            string condition;
            decimal issueNumber, lastIssueNumber = 0;
            OleDbCommand myAccessCommand = new OleDbCommand();

            // Create the dataset:
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = new OleDbConnection(strAccessConn);
            index = Convert.ToInt16(IssueGridView.Rows[0].Cells[0].Value);
            // If updating issues, delete the old ones first
            if ((string)IssueAddOrUpdateButton.Tag == "Update") {
                strAccess = "DELETE * FROM Issues where index = " + index.ToString();
                try {
                    myAccessConn.Open();
                    myAccessCommand.Connection = myAccessConn;
                    myAccessCommand.CommandText = strAccess;
                    int temp = myAccessCommand.ExecuteNonQuery();
                    if (temp > 0) {
                        Console.WriteLine("Issue Information Deleted OK");
                    } else {
                        Console.WriteLine("Error deleting issue information");
                    }
                } catch (Exception ex) {
                    Console.WriteLine("Error: Failed to delete the data from the DataBase.\n{0} - ", ex.Message);
                    myAccessConn.Close();
                    return;
                }
            } else {    // If adding issues, just open the connection
                try {
                    myAccessConn.Open();
                } catch (Exception ex) {
                    Console.WriteLine("Error: Failed to open connection to database - " + ex.Message);
                    return;
                }
            }
            // Now put in what is in the data grid for the issues
            for (int i = 0; i < Convert.ToInt16(IssueGridView.RowCount) - 1; i++) {
                index = Convert.ToInt16(IssueGridView.Rows[0].Cells[0].Value);
                issueNumber = Convert.ToDecimal(IssueGridView.Rows[i].Cells[1].Value);
                condition = Convert.ToString(IssueGridView.Rows[i].Cells[2].Value);
                retailPrice = Convert.ToDecimal(IssueGridView.Rows[i].Cells[3].Value);
                investmentValue = Convert.ToDecimal(IssueGridView.Rows[i].Cells[4].Value);
                collectionValue = Convert.ToDecimal(IssueGridView.Rows[i].Cells[5].Value);
                strAccess = @"INSERT INTO Issues ([Index], Issue, Condition, RetailPrice, InvestmentValue, CollectionValue) VALUES (" +
                    index.ToString() + "," +
                    issueNumber.ToString() + ",'" +
                    condition + "'," +
                    retailPrice.ToString() + "," +
                    investmentValue.ToString() + "," +
                    collectionValue.ToString() + ")";
                try {
                    myAccessCommand.Connection = myAccessConn;
                    myAccessCommand.CommandText = strAccess;
                    int temp = myAccessCommand.ExecuteNonQuery();
                    if (temp > 0) {
                        Console.WriteLine("Issue Information inserted OK");
                    } else {
                        Console.WriteLine("Error inserting issue information");
                    }
                } catch (Exception ex) {
                    Console.WriteLine("Error: Failed to insert the data from the DataBase.\n{0}\n" + strAccess, ex.Message);
                    MessageBox.Show("Error: Failed to insert the data from the DataBase with" + strAccess + ex.Message);
                    myAccessConn.Close();
                    return;
                }
                // Save highest number for updating title record
                if (lastIssueNumber < issueNumber) {
                    lastIssueNumber = issueNumber;
                }
                Console.WriteLine("Last Issue: " + lastIssueNumber.ToString());
            }
            TitleGridView.CurrentRow.Cells["LastIssue"].Value = lastIssueNumber;
            TitleGridView.CurrentRow.Cells["NumberOfIssues"].Value = Convert.ToInt16(IssueGridView.RowCount) - 1; // Number of issues

            myAccessConn.Close();
            // Default the action back to Update so that Add is not the default.
            IssueAddOrUpdateButton.Tag = "Update";
            // TODO: IssueGridView.Sort(IssueGridView.Columns[1], ListSortDirection.Ascending);  // Resort the view to show update
        }


        /// <summary>
        /// Button Click handler for Add button for latest issues based on Selected Title 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddLatestButton_Click(object sender, EventArgs e) {
            int index = 0, numberOfIssues = 0, numberOfIssuesToAdd = 0;
            Title localTitle = new Title();
            Issue localIssue = new Issue();
            decimal retailPrice = 0;

            // Get index from currently selected title
            try {
                index = Convert.ToInt16(TitleGridView.SelectedRows[0].Cells[0].Value);

                numberOfIssues = Utilities.fillCurrentIssues(index);    // Fill in currentIssues array from database
                localIssue = Utilities.currentIssues[numberOfIssues - 1];   // Get last issue for this title
                retailPrice = localIssue.retailPrice;   // Get last issue's retail price
                // Avoid processing a null entry
                numberOfIssuesToAdd = Convert.ToInt16(numberToAddUpDown.Value);
                
                // Add the issues to the end of the grid view
                for (int i = 0; i < numberOfIssuesToAdd; i++) {
                    IssueGridView.Rows.Add(index, ++localIssue.issueNumber, localIssue.condition, localIssue.retailPrice,
                        localIssue.investmentValue, localIssue.collectionValue);
                }
                // 'Push' the Update Issues button to get the issues into their table - also updates the title information
                IssueAddOrUpdateButton.Tag = "Update";  // Ensure we do a replacement
                IssueAddOrUpdateButton_Click(sender, e); // Force the issues into the database
                numberToAddUpDown.Value = 1;    // reset the number to add to avoid left over value adding issues if button hit again
            } catch (Exception ex) {
                Console.WriteLine("Exception in click of Add button - Selected row not valid " + ex.Message);
            }
        }

        /// <summary>
        /// Handle the user deleting the row event. Need to get the index for the row selected for deletion and 
        /// delete the issues that correspond to that index.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void TitleGridView_UserDeletingRow(object sender, System.Windows.Forms.DataGridViewRowCancelEventArgs e) {
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Utilities.comicDatabaseFileName;
            string strAccess = "";
            int index = 0;
            OleDbCommand myAccessCommand = new OleDbCommand();

            try {
                index = Convert.ToInt16(e.Row.Cells[0].Value);  // Get the index value for the database lookup of the issues to delete
                Utilities.missingTitleIndex[index] = -1;        // Mark as free
                OleDbConnection myAccessConn = new OleDbConnection(strAccessConn);
                // Delete the issues that match the index of the title being deleted
                strAccess = "DELETE * FROM Issues where index = " + index.ToString();
                try {
                    myAccessConn.Open();
                    myAccessCommand.Connection = myAccessConn;
                    myAccessCommand.CommandText = strAccess;
                    int temp = myAccessCommand.ExecuteNonQuery();
                    if (temp > 0) {
                        Console.WriteLine("Issue Information Deleted OK");
                    } else {
                        Console.WriteLine("Error deleting issue information");
                    }
                } catch (Exception ex) {
                    Console.WriteLine("Error: Failed to delete the data from the DataBase.\n{0}", ex.Message);
                    myAccessConn.Close();
                    return;
                }
            } catch (Exception ex) {
                Console.WriteLine("Error: SelectedRows of TitleGridView out of range -- " + ex.Message);
            }

        }

        /// <summary>
        /// User has entered a value into the last row of the Issue Grid. Copy the Index value for this Title from
        /// the first Row into the Index column for the new Row.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void IssueGridView_UserAddedRow(object sender, DataGridViewRowEventArgs e) {
            int index;
            index = Convert.ToInt16(IssueGridView.Rows[0].Cells[0].Value);  // Get current index for these issue
            IssueGridView.CurrentRow.Cells[0].Value = index;
        }

        /// <summary>
        /// User has entered a value into the last row of the Title Grid. Set the Index Value for this Title from this new
        /// row into the Index column for the new Row.
        /// Then need to create a dummy entry in the Issue Grid for the first issue of the new title.
        /// Also handle the missingTitleIndex array update - used to keep track of where to put new titles to avoid duplicate entries.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TitleGridView_UserAddedRow(object sender, DataGridViewRowEventArgs e) {
            int index = 0;
            // Find first missing title record and use that to enter the new title information.
            for (index = 1; index < Utilities.maxTitleLimit; index++) {
                if (Utilities.missingTitleIndex[index] == -1) {
                    Console.WriteLine("Max Title Index found OK");
                    TitleGridView.CurrentRow.Cells[0].Value = index;
                    TitleGridView.CurrentRow.Cells[9].Value = 0;    // Set update year to 0 to avoid a null value
                    IssueAddOrUpdateButton.Tag = "Add";
                    IssueGridView.Rows[0].Cells[0].Value = index;
                    Utilities.missingTitleIndex[index] = index; // Mark as used
                    //TitleGridView.Sort(TitleGridView.Columns[0], ListSortDirection.Ascending);  // Sort the Title grid by the Index (Column 0)
                    break;
                }
            }

        }

        /// <summary>
        /// Set visible row to first title matching tag value of label
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Alabel_Click(object sender, EventArgs e) {
            Label mLabel = (Label) sender;
            // Get tag value from label and use it to look up first matching title in the TitleGridView
            foreach (DataGridViewRow row in TitleGridView.Rows) {
                if (row.Cells["Title"].Value.ToString()[0] == (mLabel.Tag.ToString().ToUpper()[0])) {
                    TitleGridView.FirstDisplayedScrollingRowIndex = row.Index;
                    break;  // Done, can leave loop
                }
            }
        }

        /// <summary>
        /// If the user sorts the titles or otherwise changes information, then re-paint the background based on the incomplete status
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TitleGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            // If the last issue is less than the number of issues, then the title is missing some issues.
            // Set the background to a pink value.
            int numberofissues = 0;
            int lastissue = 0;
            try {
                numberofissues = Convert.ToInt16(TitleGridView.Rows[e.RowIndex].Cells["NumberOfIssues"].Value);
                lastissue = Convert.ToInt16(TitleGridView.Rows[e.RowIndex].Cells["LastIssue"].Value);
                if (numberofissues < lastissue) {
                    TitleGridView.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightPink;
                }
            } catch (Exception ex) {
                Console.WriteLine("Exception handling cell formatting in row " + e.RowIndex.ToString() + " " + ex.Message);
            }
        }

        private void TitleGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e) {
            DataGridView mGrid = (DataGridView) sender;
            // Turn off callback since it will recurse if we change the value here.
            TitleGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.TitleGridView_CellValueChanged);

            if (mGrid.CurrentCellAddress.X == 4) {
                // Current check box changed
                if (Convert.ToBoolean(mGrid.CurrentCell.Value)) {
                    mGrid.CurrentRow.Cells["Type"].Value = Convert.ToInt16(Convert.ToInt16(mGrid.CurrentRow.Cells["Type"].Value) | 1);
                } else {
                    if (Convert.ToBoolean(mGrid.CurrentRow.Cells["Complete"].Value)) {
                        mGrid.CurrentRow.Cells["Type"].Value = 2;
                    } else {
                        mGrid.CurrentRow.Cells["Type"].Value = 0;
                    }
                }
            } else if (mGrid.CurrentCellAddress.X == 5) {
                // Complete check box changed
                if (Convert.ToBoolean(mGrid.CurrentCell.Value)) {
                    mGrid.CurrentRow.Cells["Type"].Value = Convert.ToInt16(mGrid.CurrentRow.Cells["Type"].Value) | 2;
                } else {
                    if (Convert.ToBoolean(mGrid.CurrentRow.Cells["Current"].Value)) {
                        mGrid.CurrentRow.Cells["Type"].Value = 1;
                    } else {
                        mGrid.CurrentRow.Cells["Type"].Value = 0;
                    }
                }
            } if (mGrid.CurrentCellAddress.X == 3) {
                // Type value changed - set check boxes for Current and Complete
                if (Convert.ToInt16(mGrid.CurrentRow.Cells["Type"].Value) == 0) {
                    // Not current, nor complete
                    mGrid.CurrentRow.Cells["Current"].Value = false;
                    mGrid.CurrentRow.Cells["Complete"].Value = false;
                } else if (Convert.ToInt16(mGrid.CurrentRow.Cells["Type"].Value) == 1) {
                    // Current, not complete
                    mGrid.CurrentRow.Cells["Current"].Value = true;
                    mGrid.CurrentRow.Cells["Complete"].Value = false;
                } else if (Convert.ToInt16(mGrid.CurrentRow.Cells["Type"].Value) == 2) {
                    // Complete, not current
                    mGrid.CurrentRow.Cells["Current"].Value = false;
                    mGrid.CurrentRow.Cells["Complete"].Value = true;
                } else if (Convert.ToInt16(mGrid.CurrentRow.Cells["Type"].Value) == 3) {
                    // Complet and Current
                    mGrid.CurrentRow.Cells["Current"].Value = true;
                    mGrid.CurrentRow.Cells["Complete"].Value = true;
                } else {
                    Console.WriteLine("Invalide Type value entered: " + mGrid.CurrentRow.Cells["Type"].Value.ToString());
                }
            }
            // put callback back for next use.
            TitleGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.TitleGridView_CellValueChanged);
        }

        /// <summary>
        /// Connected to Full Collection Summary Export menu item.
        /// Create a CSV file with Title information and Retail, Investment and Collection values.
        /// Can be used to determine most valuable comic runs in collection.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void fullCollectionSummaryExportMenuItem_Click(object sender, EventArgs e) {
            StreamWriter SW;
            SW = File.CreateText(Utilities.outputFileName);
            Title localTitle = new Title();
            this.Cursor = Cursors.WaitCursor;
            int numberOfTitles = TitleGridView.RowCount - 1;
            comicStatusLabel.Text = "Printing Title File Report";
            string outputLine = "";
            decimal totalRetailValue = 0.0M, totalInvestmentValue = 0.0M, totalCollectionValue = 0.0M;
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Utilities.comicDatabaseFileName;
            string strAccess = "";

            int index;

            // Sort the TitleGridView by the Title name before printing
            TitleGridView.Sort(TitleGridView.Columns[1], ListSortDirection.Ascending);  // Sort the Title grid by the Title Name (Column 1)

            OleDbConnection myAccessConn = null;
            try {
                myAccessConn = new OleDbConnection(strAccessConn);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }

            DataSet myDataSet = new DataSet();

            try {
                myAccessConn.Open();
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            } 

            outputLine = localTitle.CSVHeaderOut + ",Total Retail Value," + "Total Investment Value, " + "Total Collection Value";

            SW.WriteLine(outputLine);
            for (int title = 0; title < numberOfTitles; title++) {
                localTitle = Utilities.ReadTitle(TitleGridView.Rows, title);
                comicStatusLabel.Text = "Processing: " + localTitle.title.ToString();
//                Console.WriteLine("Processing: " + localTitle.title.ToString());
                index = localTitle.index;
                strAccess = "SELECT * FROM Issues where index = " + index.ToString() + " ORDER BY Issue";
                OleDbCommand myAccessCommand = new OleDbCommand(strAccess, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(myDataSet);
                computeTitleValues(localTitle, myDataSet, ref totalRetailValue, ref totalInvestmentValue, ref totalCollectionValue);
                outputLine = localTitle.CSVOut + "," + totalRetailValue.ToString() + "," + totalInvestmentValue.ToString() + "," + totalCollectionValue.ToString();
                SW.WriteLine(outputLine);
                totalCollectionValue = totalInvestmentValue = totalRetailValue = 0.0M;
                myDataSet.Reset();
                Application.DoEvents();
            }
            myAccessConn.Close();
            SW.Close();

            this.Cursor = Cursors.Arrow;
            comicStatusLabel.Text = "Done";
        }

        /// <summary>
        /// Output a report on the currently selected title. List the title and its information along with all 
        /// the issues for this title to the current output file.
        /// </summary>
        /// <param name="localTitle">Currently selected title to report.</param>
        private void computeTitleValues(Title localTitle, DataSet myDataSet, ref decimal totalRetailValue, ref decimal totalInvestmentValue, ref decimal totalCollectionValue) {
            int numberOfIssues = 0;
            int index = 0;
            Issue localIssue = new Issue();

            this.Cursor = Cursors.WaitCursor;
            comicStatusLabel.Text = "Printing Title Report: " + localTitle.title;

            // Read all the issues for this title and generate the report information for them
            index = localTitle.index;

            try {
                DataRowCollection dra = myDataSet.Tables["Table"].Rows;
                localIssue = Utilities.ReadIssue(dra[0]);
                foreach (DataRow dr in dra) {
                    localIssue = Utilities.ReadIssue(dr);
                    if (localIssue.collectionValue == 0.0M) {
                        totalCollectionValue += localIssue.retailPrice;
                    } else {
                        totalCollectionValue += localIssue.collectionValue;
                    }
                    if (localIssue.investmentValue == 0.0M) {
                        totalInvestmentValue += localIssue.retailPrice;
                    } else {
                        totalInvestmentValue += localIssue.investmentValue;
                    }
                    totalRetailValue += Convert.ToDecimal(localIssue.retailPrice);
                    numberOfIssues++;

                }
            } catch (Exception ex) {
                Console.WriteLine("Exception: " + ex.Message + " processing issue information");
            }

            this.Cursor = Cursors.Arrow;
            comicStatusLabel.Text = "Done";
        }

        /// <summary>
        /// Set the Collection Value and Investment Value of each issue in the database to the value of the Retail Price if they are 0.00.
        /// This will allow straight summations to be used on these values instead of using the retail prices as these values if they are 0.00
        /// as is done now.
        /// This is pretty much a one time operation but can be used after a number of new issues have been added to set those also.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void setCVAndIVToolStripMenuItem_Click(object sender, EventArgs e) {

        }

        private void numberToAddUpDown_ValueChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// Alphabatize the titles and select the current ones for easier updating each month
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void selectCurrentToolStripMenuItem_Click(object sender, EventArgs e) {
            TitleGridView.Sort(TitleGridView.Columns[1], ListSortDirection.Ascending);  // Sort the Title grid by the Title Name (Column 1)
            TitleGridView.Sort(TitleGridView.Columns[4], ListSortDirection.Descending); // Then select the current ones            
        }
    }
}