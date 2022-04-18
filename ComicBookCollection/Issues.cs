using System;
using System.Collections.Generic;
using System.Text;

namespace ComicCollector {
    class Issue {
        private int mLink;
        private decimal mIssueNumber;
        private decimal mRetailPrice;
        private decimal mInvestmentValue;
        private decimal mCollectionValue;
        private string mCondition;


        public int link {
            get { return mLink; }
            set { mLink = value; }
        }

        public decimal issueNumber {
            get { return mIssueNumber; }
            set { mIssueNumber = value; }
        }

        public decimal retailPrice {
            get { return mRetailPrice; }
            set { mRetailPrice = value; }
        }

        public decimal investmentValue {
            get { return mInvestmentValue; }
            set { mInvestmentValue = value; }
        }

        public decimal collectionValue {
            get { return mCollectionValue; }
            set { mCollectionValue = value; }
        }

        public string condition {
            get { return mCondition; }
            set { mCondition = value; }
        }



        public string report {
            get {
                return "Issue: " + issueNumber.ToString() +
                    " (RP " + retailPrice.ToString() + ") " +
                    " (INV " + investmentValue.ToString() + ") " +
                    "[COL " + collectionValue.ToString() + "]" +
                    "<Grade " + condition.ToString() + ">";
            }
        }

        public string printout {
            get {
                return issueNumber.ToString() +
                    " (" + investmentValue.ToString() + ")" +
                    " [" + collectionValue.ToString() + "]" +
                    " <" + condition.ToString() + ">";
            }
        }

        public string diagnosticOut {
            get {
                return "Issue: " + issueNumber.ToString() +
                    " Link: " + link.ToString() +
                    " Retail: " + retailPrice.ToString() +
                    " Invest: " + investmentValue.ToString() +
                    " Collection: " + collectionValue.ToString() +
                    " Condition: " + condition.ToString() + "\r\n";
            }
        }

        public string CSVOut(int record) {
            return record.ToString() + "," + issueNumber.ToString() + "," + condition.ToString() + "," + retailPrice.ToString() + "," +
                investmentValue.ToString() + "," + collectionValue.ToString();
        }

        public string CSVOut() {
            return link.ToString() + "," + issueNumber.ToString() + "," + condition.ToString() + "," + retailPrice.ToString() + "," +
                investmentValue.ToString() + "," + collectionValue.ToString();
        }

        public string CSVHeaderOut {
            get {
                return "Index" + "," + "Issue" + "," + "Condition" + "," + "RetailPrice" + "," + "InvestmentValue" + "," + "CollectionValue";
            }
        }
    }


}