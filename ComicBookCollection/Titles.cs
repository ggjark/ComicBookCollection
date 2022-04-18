using System;
using System.Collections.Generic;
using System.Text;

namespace ComicCollector {
    class Title {
        private int mIndex;
        private string mTitle;
        private string mPublisher;
        private int mType;
        private int mNumberOfIssues;
        private int mLastIssue;
        private bool mCurrent;
        private bool mComplete;
        private bool mLegacy;
        private int mUpdateYear;

        public int index {
            get { return mIndex;}
            set { mIndex = value; }
        }

        public string title {
            get { return mTitle; }
            set { mTitle = value; }
        }

        public string titleQuoted {
            get { return "\"" + mTitle.ToString() + "\""; }
        }

        public string publisher {
            get { return mPublisher; }
            set { mPublisher = value; }
        }

        public int type {
            get { return mType; }
            set { mType = value; }
        }

        public int numberOfIssues {
            get { return mNumberOfIssues; }
            set { mNumberOfIssues = value; }
        }

        public int lastIssue {
            get { return mLastIssue; }
            set { mLastIssue = value; }
        }

        public bool current {
            get { return mCurrent; }
            set { mCurrent = value; }
        }

        public bool complete {
            get { return mComplete; }
            set {mComplete = value; }
        }

        public bool legacy {
            get { return mLegacy;  }
            set { mLegacy = value; }
        }

        public int updateyear {
            get { return mUpdateYear; }
            set { mUpdateYear = value; }
        }

        public string diagnosticOut {
            get {
                return "Title: " + title +
                    "\r\n    Publisher: " + publisher.ToString() +
                    " Type: " + type.ToString() +
                    " Number of Issues: " + numberOfIssues.ToString() +
                    " Last Issue: " + lastIssue.ToString() +
                    " Current: " + current.ToString() +
                    " Complete: " + complete.ToString() +
                    " Legacy: " + legacy.ToString() +
                    " Update Year: " + updateyear.ToString() +
                    "\r\n\r\n";
            }
        }

        public string CSVOut {
            get {
                return index.ToString() + "," + titleQuoted + "," + publisher.ToString() + "," + type.ToString() + "," + 
                    current.ToString() + "," + complete.ToString() + "," + numberOfIssues.ToString() + "," + 
                    lastIssue.ToString() + "," + legacy.ToString() + "," + updateyear.ToString();
            }
        }

        public string CSVHeaderOut {
            get {
                return "Index," + "Title," + "Publisher," + "Type," + "Current," + "Complete," + "Number of Issues," + "Last Issue," + "Legacy," + "Update Year";
            }
        }
    
    }
}
