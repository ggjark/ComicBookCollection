using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;



namespace ComicCollector {
    class Utilities {

        const string qualityTypes = "P FAG VGFIVFNMM PM";
        public const int titleSectionLength = 35;
        public const int titleRecordLength = 49;
        public const int issueRecordLength = 18;
        public const int portraitLineLength = 80;
        public const int landscapeLineLength = 132;
        public const int maxIssuesPerTitle = 2000;
        public const int maxTitleLimit = 10000;
        public static string comicDatabaseFileName = "ComicDataBase.accdb";
        public static string verifyOutFileName = "verify.txt";
        public static string latestOutFileName = "latest.txt";
        public static string outputFileName = "output.txt";
        public static string CSVFileForNames = "names.csv";
        public static string CSVFileForIssues = "issues.csv";
        /// <summary>
        /// Holds the data on the current issues for the currently selected title
        /// </summary>
        public static Issue[] currentIssues = new Issue[Utilities.maxIssuesPerTitle];
        /// <summary>
        /// Holds the missing Titles index values from when Titles are deleted
        /// </summary>
        public static int[] missingTitleIndex = new int[maxTitleLimit];



        /// <summary>
        /// Convert Title data from the Database DataRow as set in the Title table
        /// </summary>
        /// <param name="DataRow dr"></param>
        /// <returns>Title Record</returns>
        public static Title ReadTitle(DataRow dr) {
            Title localTitle = new Title();
            localTitle.index = Convert.ToInt16(dr[0]);
            localTitle.title = Convert.ToString(dr[1]);
            localTitle.publisher = Convert.ToString(dr[2]);
            localTitle.type = Convert.ToInt16(dr[3]);
            localTitle.current = Convert.ToBoolean(dr[4]);
            localTitle.complete = Convert.ToBoolean(dr[5]);
            localTitle.numberOfIssues = Convert.ToInt16(dr[6]);
            localTitle.lastIssue = Convert.ToInt16(dr[7]);
            localTitle.legacy = Convert.ToBoolean(dr[8]);
            localTitle.updateyear = Convert.ToInt16(dr[9]);
            return localTitle;
        }

        public static Title ReadTitle(DataGridViewRowCollection row, int title) {
            Title localTitle = new Title();
            localTitle.index = Convert.ToInt16(row[title].Cells[0].Value);
            localTitle.title = Convert.ToString(row[title].Cells[1].Value);
            localTitle.publisher = Convert.ToString(row[title].Cells[2].Value);
            localTitle.type = Convert.ToInt16(row[title].Cells[3].Value);
            localTitle.current = Convert.ToBoolean(row[title].Cells[4].Value);
            localTitle.complete = Convert.ToBoolean(row[title].Cells[5].Value);
            localTitle.numberOfIssues = Convert.ToInt16(row[title].Cells[6].Value);
            localTitle.lastIssue = Convert.ToInt16(row[title].Cells[7].Value);
            localTitle.legacy = Convert.ToBoolean(row[title].Cells[8].Value);
            localTitle.updateyear = Convert.ToInt16(row[title].Cells[9].Value);
            return localTitle;
        }

        public static Issue ReadIssue(DataRow dr) {
            Issue localIssue = new Issue();
            localIssue.link = Convert.ToInt16(dr[0]);
            localIssue.issueNumber = Convert.ToDecimal(dr[1]);
            localIssue.condition = Convert.ToString(dr[2]);
            localIssue.retailPrice = Convert.ToDecimal(dr[3]);
            localIssue.investmentValue = Convert.ToDecimal(dr[4]);
            localIssue.collectionValue = Convert.ToDecimal(dr[5]);
            return localIssue;
        }


        /// <summary>
        /// Read set of Issues from database based on index supplied into global currentIssues array
        /// </summary>
        /// <param name="index">Database table index to use for SQL call</param>
        /// <returns>number of issue read, -1 on failure</returns>
        public static int fillCurrentIssues(int index) {
            string strAccessSelect = "";
            // Create the dataset and add the Categories table to it:
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            string strAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Utilities.comicDatabaseFileName;
            Issue localIssue = new Issue();
            int i;

            strAccessSelect = "SELECT * FROM Issues where index = " + index.ToString() + " ORDER BY Issue";
            try {
                myAccessConn = new OleDbConnection(strAccessConn);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return -1;
            }
            try {
                myAccessConn.Open();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(myDataSet);
            } catch (Exception ex) {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return -1;
            } finally {
                myAccessConn.Close();
            }

            i = 0;
            DataRowCollection dra = myDataSet.Tables["Table"].Rows;
            
            foreach (DataRow dr in dra) {
                localIssue = Utilities.ReadIssue(dr);
                Utilities.currentIssues[i++] = localIssue;
            }
            return i;

        }

    }
}
