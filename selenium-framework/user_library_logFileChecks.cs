using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Windows.Controls.Primitives;

namespace selenium_framework
{
    // Common library for frequently used functionality for reuse across scripts
    public partial class Selenium_Framework
    {
        // Check the student comment log file
        public (bool, string, string) checkCommentImportLogFile(string logFileLocation, bool wasTestrun)
        {
            debug(TRACE, "checkCommentImportLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int numRecordsProcessed = 0;
            int numTotalRejects = 0;
            int numTotalInserts = 0;
            int numTotalErrors = 0;
            string[] tempSplit = null;

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Total Records Processed"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numRecordsProcessed = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numRecordsProcessed = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Total Rejects"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalRejects = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalRejects = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Total Inserts"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalInserts = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalInserts = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Total Errors"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalErrors = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalErrors = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
            }

            // error validation
            if (numTotalErrors >= 1) // need an error file to check how the errors look like... 
            {
                message = "Test failed :" +
                    "\nNumber of records processed: " + numRecordsProcessed +
                   "\nNumber of records rejected: " + numTotalRejects +
                   "\nNumber of records errors: " + numTotalErrors +
                    "\nNumber of records inserted: " + numTotalInserts;
                result = false;
                return (result, typeOfFile, message);

            }
            else
            {
                message = "Test Passed :" +
                   "\nNumber of records processed: " + numRecordsProcessed +
                   "\nNumber of records rejected: " + numTotalRejects +
                   "\nNumber of records errors: " + numTotalErrors +
                   "\nNumber of records inserted: " + numTotalInserts;
                result = true;
                return (result, typeOfFile, message);
            }

        }

        public (bool, string, string) checkRecordsProcessedLogFile(string logFileLocation)
        {
            debug(TRACE, "checkRecordsProcessedLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int recsToProcess = 0;
            int recsProcessed = 0;
            int recsFailed = 0;
            int recsAltered = 0;
            int recsNotAltered = 0;
            string[] tempSplit = null;

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Records to Process"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    recsToProcess = int.Parse(tempSplit[tempSplit.Length - 1]);
                }
                else if (line.StartsWith("Records Processed"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    recsProcessed = int.Parse(tempSplit[tempSplit.Length - 1]);
                }
                else if (line.StartsWith("Records Failed"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    recsFailed = int.Parse(tempSplit[tempSplit.Length - 1]);
                }
                else if (line.StartsWith("Records Altered"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    recsAltered = int.Parse(tempSplit[tempSplit.Length - 1]);
                }
                else if (line.StartsWith("Records Not Altered"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    recsNotAltered = int.Parse(tempSplit[tempSplit.Length - 1]);
                }
            }

            // error validation
            if (recsFailed >= 1)
            {
                message = "Test failed :" +
                    "\nNumber of records processed: " + recsProcessed +
                    "\nNumber of files altered: " + recsAltered +
                    "\nNumber of files failed: " + recsFailed;
                result = false;
                return (result, typeOfFile, message);

            }
            else
            {
                message = "Test Passed :" +
                       "\nNumber of records processed: " + recsProcessed +
                       "\nNumber of files altered: " + recsAltered +
                       "\nNumber of files failed: " + recsFailed;
                result = false;
                return (result, typeOfFile, message);
            }

        }

        // utility method for generateTACFile
        public int returnHeaderIndex(string[] headers, string textToFind)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                if (headers[i] == textToFind)
                {
                    return i;
                }
            }
            return 999; // not found
        }

        // Generates a basic TAC import file for testing
        // no international students
        public string generateTACFile(int studentCount, string outFile)
        {
            debug(TRACE, "generateTACFile: start");
            debug(TRACE, "generateTACFile: # of students to generate: " + studentCount);
            debug(TRACE, "generateTACFile: output file location: " + outFile);

            // output file values
            string offRound = RandomNumber(4);
            string headerText = "FORMAT OFFER STANDARD";

            // data files
            string baseFile = dirName + @"\TACData\input_2023.dat";
            string dictionaryFile = dirName + @"\TACData\dictionary.txt";
            string courseFile = dirName + @"\TACData\newug-spring.csv"; // orig ug.csv newug-aw newug-aw

            DateTime now = DateTime.Now;
            string buildTACAppNumber = now.Hour.ToString() + now.Minute.ToString() + "00000";
            int giTACAppNumber = int.Parse(buildTACAppNumber);

            string gsOfferDate = now.ToString("dd-MMM-yy");
            string gsLapseDate = now.AddMonths(1).ToString("dd-MMM-yy");

            // Load dictionary file
            string[] dictionaryEntries = File.ReadAllLines(dictionaryFile);

            // Load course file
            string[] courseEntries = File.ReadAllLines(courseFile);

            List<string> newlines = new List<string>();
            string[] headerSplit = null;

            Random rnd = new Random();

            using (StreamReader reader = new StreamReader(baseFile))
            {
                string line;
                int counter = 1;

                while ((line = reader.ReadLine()) != null && counter <= studentCount)
                {
                    if (counter == 1)
                    {
                        newlines.Add(headerText);
                    }
                    if (counter == 2)
                    {
                        // split this line to use as index
                        newlines.Add(line);
                        headerSplit = line.Split(new char[] { ',' });
                    }

                    if (counter >= 3)
                    {

                        // get random name and course
                        int myRandom = rnd.Next(dictionaryEntries.Length);
                        string lsName = dictionaryEntries[myRandom];

                        myRandom = rnd.Next(courseEntries.Length);
                        string lsCourse = courseEntries[myRandom];

                        string[] lineSplit = line.Split(new char[] { ',' });

                        // update the application number and offer round
                        lineSplit[returnHeaderIndex(headerSplit, "Stu_alt_id")] = (giTACAppNumber++).ToString();
                        lineSplit[returnHeaderIndex(headerSplit, "Offer_round")] = offRound;

                        // create new application / lapse dates
                        lineSplit[returnHeaderIndex(headerSplit, "Offer_dt")] = gsOfferDate;
                        lineSplit[returnHeaderIndex(headerSplit, "Lapse_dt")] = gsLapseDate;

                        // blank the student number
                        lineSplit[returnHeaderIndex(headerSplit, "Stu_id")] = "";

                        // blank out CHESSN
                        lineSplit[returnHeaderIndex(headerSplit, "Gsn_no")] = "";

                        // Generate names
                        lineSplit[returnHeaderIndex(headerSplit, "Family_nm")] = lsName + RandomString(4);
                        lineSplit[returnHeaderIndex(headerSplit, "Gvn_nm")] = lsName + RandomString(4);
                        lineSplit[returnHeaderIndex(headerSplit, "Oth_nm")] = lsName + RandomString(4);
                        lineSplit[returnHeaderIndex(headerSplit, "Prev_family_nm_1")] = lsName + RandomString(4);
                        lineSplit[returnHeaderIndex(headerSplit, "Prev_family_nm_2")] = lsName + RandomString(4);

                        // email address
                        lineSplit[returnHeaderIndex(headerSplit, "Email_type_addr_1")] = lsName + RandomString(2) + "@" + RandomString(10) + ".foo";

                        // build date of birth
                        int randNumber = rnd.Next(24, 30);
                        string newDOB = now.AddYears(-randNumber).ToString("dd-MMM-yy");
                        lineSplit[returnHeaderIndex(headerSplit, "Birth_dt")] = newDOB;

                        // assign the final chosen course to the student
                        lineSplit[returnHeaderIndex(headerSplit, "Admsn_ctr_crs_cd")] = lsCourse;

                        string newLine = "";
                        for (int i = 0; i < lineSplit.Length; i++)
                        {
                            if (i != lineSplit.Length - 1)
                            {
                                newLine = newLine + lineSplit[i] + ",";
                            }
                            else
                            {
                                newLine = newLine + lineSplit[i];
                            }
                        }

                        // need to rebuild the line before this
                        newlines.Add(newLine);
                    }
                    //Console.WriteLine(line);
                    counter++;
                }
                using (TextWriter tw = new StreamWriter(outFile))
                {
                    foreach (var item in newlines)
                    {
                        tw.WriteLine(item);
                    }
                }
                sleep(1);

            }
            return offRound;
        }

        // Check the log files for both sanctions amd TAC import
        public (bool, string, string) checkTACLogFile(string logFileLocation, bool wasTestrun)
        {
            debug(TRACE, "checkTACLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int numLinesFiles = 0;
            int numDetailLinesRead = 0;
            int numDetailLinesAnomalies = 0;
            int numRecordsCreated = 0;
            int numRecordsUpdated = 0;
            string[] tempSplit = null;

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Number of lines in file"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numLinesFiles = int.Parse(tempSplit[tempSplit.Length - 2]);
                    }
                    else
                    {
                        numLinesFiles = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Number of detail lines read"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numDetailLinesRead = int.Parse(tempSplit[tempSplit.Length - 2]);
                    }
                    else
                    {
                        numDetailLinesRead = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Number of detail lines with anomalies"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numDetailLinesAnomalies = int.Parse(tempSplit[tempSplit.Length - 2]);
                    }
                    else
                    {
                        numDetailLinesAnomalies = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Number of records to be created"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numRecordsCreated = int.Parse(tempSplit[tempSplit.Length - 3]);
                    }
                    else
                    {
                        numRecordsCreated = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Number of records to be updated"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numRecordsUpdated = int.Parse(tempSplit[tempSplit.Length - 3]);
                    }
                    else
                    {
                        numRecordsUpdated = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
            }

            // error validation
            //if (recsFailed >= 1) // need an error file to check how the errors look like... 
            //{
            //    message = "Test failed :" +
            //        "\nNumber of records created: " + numRecordsCreated +
            //       "\nNumber of files updated: " + numRecordsUpdated +
            //        "\nNumber of files anomalies: " + numDetailLinesAnomalies;
            //    result = false;
            //    return (result, typeOfFile, message);

            //}
            //else
            //{
            message = "Test Passed :" +
                   "\nNumber of records created: " + numRecordsCreated +
                   "\nNumber of files updated: " + numRecordsUpdated +
                   "\nNumber of files anomalies: " + numDetailLinesAnomalies;
            result = false;
            return (result, typeOfFile, message);
            //}

        }

        // Check the TAC (UAC) load log file
        public (bool, string, string) checkTACLoadLogFile(string logFileLocation, bool wasTestrun)
        {
            debug(TRACE, "checkTACLoadLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int numRecordsProcessed = 0;
            int numTotalRejects = 0;
            int numTotalInserts = 0;
            string[] tempSplit = null;

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Total Records Processed"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numRecordsProcessed = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numRecordsProcessed = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Total Rejects"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalRejects = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalRejects = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Total Inserts"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalInserts = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalInserts = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
            }

            // error validation
            if (numTotalRejects >= 1) // need an error file to check how the errors look like... 
            {
                message = "Test failed :" +
                    "\nNumber of records processed: " + numRecordsProcessed +
                   "\nNumber of records rejected: " + numTotalRejects +
                    "\nNumber of records inserted: " + numTotalInserts;
                result = false;
                return (result, typeOfFile, message);

            }
            else
            {
                message = "Test Passed :" +
                   "\nNumber of records processed: " + numRecordsProcessed +
                   "\nNumber of records rejected: " + numTotalRejects +
                   "\nNumber of records inserted: " + numTotalInserts;
                result = true;
                return (result, typeOfFile, message);
            }

        }

        public (bool, string, string) checkBulkCheckingBestowalLogFile(string logFileLocation, bool wasTestrun)
        {
            debug(TRACE, "checkBulkCheckingBestowalLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int numStudentSetsProcessed = 0;
            int numErrors = 0;
            int numWarnings = 0;
            int numInfoMessages = 0;
            string[] tempSplit = null;

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Student Sets Processed"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numStudentSetsProcessed = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numStudentSetsProcessed = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Errors"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numErrors = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numErrors = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Warnings"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numWarnings = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numWarnings = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Information Messages"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numInfoMessages = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numInfoMessages = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
            }

            // error validation
            if (numErrors >= 1) // need an error file to check how the errors look like... 
            {
                message = "Test failed :" +
                    "\nNumber of Student Sets Processed: " + numStudentSetsProcessed +
                   "\nNumber of errors: " + numErrors +
                    "\nNumber of warnings: " + numWarnings;
                result = false;
                return (result, typeOfFile, message);

            }
            else
            {
                message = "Test Passed :" +
                   "\nNumber of Student Sets Processed: " + numStudentSetsProcessed +
                   "\nNumber of errors: " + numErrors +
                    "\nNumber of warnings: " + numWarnings;
                result = true;
                return (result, typeOfFile, message);
            }

        }

        // Check Formula Constant Roll Over Log File
        public bool checkFormulaConstantRollOverLogFileForErrors(string logFileLocation)
        {
            debug(TRACE, "checkFormulaConstantRollOverLogFile: start");

            bool lineFound = false;
            int noOfErrors = 0;
            string[] inputLines = File.ReadAllLines(logFileLocation);
            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("No. of Errors"))
                {
                    noOfErrors = int.Parse(line.Split(':')[1].Trim());
                    Console.WriteLine("!!!noOfErrors:" + noOfErrors);
                    lineFound = true;
                    break;
                }
            }

            if (lineFound == false)
            {
                debug(ERROR, "No of errors line not found in the log file");
            }

            if (noOfErrors == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool checkRequisiteWaiverLogFile(string requisiteWaiverLogsFileFullPath)
        {
            debug(TRACE, "checkRequisiteWaiverLogFile: start");

            bool status = true;
            int numberOfRecordsCreated = 0;
            int extractedNumberOfRecordsCreated = 0;
            string[] inputLines = File.ReadAllLines(requisiteWaiverLogsFileFullPath);
            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Information: Requisite Waiver"))
                {
                    if (line.Trim().EndsWith("has been created."))
                    {
                        numberOfRecordsCreated++;

                    }
                    else
                    {
                        status = false;
                        return status;
                    }

                }

                if (line.StartsWith("Number of records to be created."))
                {
                    extractedNumberOfRecordsCreated = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[6]);
                }
            }
            Console.WriteLine("numberOfRecordsCreated: " + numberOfRecordsCreated);
            Console.WriteLine("extractedNumberOfRecordsCreated: " + extractedNumberOfRecordsCreated);
            if (extractedNumberOfRecordsCreated != numberOfRecordsCreated)
            {
                status = false;

            }
            return status;
        }

        // Check Disbursee Fee Log File
        public bool checkDisburseeFeeLogFile(string disburseFeeLogsFileFullPath)
        {
            debug(TRACE, "checkDisburseeFeeLogFile: start");

            bool status = false;
            string[] inputLines = File.ReadAllLines(disburseFeeLogsFileFullPath);
            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Disburse Fees completed SUCCESSFULLY"))
                {
                    status = true;
                    break;
                }
            }
            return status;
        }

        // extract data from Disburse Fee job log message
        public (bool, string, bool, string) extractDataFromDisburseFeeJobLogMessage(string disburseFeeJobLogMessage)
        {
            debug(TRACE, "extractDataFromDisburseFeeJobLogMessage: start");

            bool foundJobNumber = false;
            string jobNumber = "";
            bool foundLogFileName = false;
            string logFileName = "";

            // 1. extract job number
            // define the regular expression pattern
            string patternForJobNumber = @"Starting job (\d+) \(Disburse Fees\)";

            // use Regex.Match to find the match in the text
            Match matchForJobNumber = Regex.Match(disburseFeeJobLogMessage, patternForJobNumber);

            if (matchForJobNumber.Success)
            {
                jobNumber = matchForJobNumber.Groups[1].Value.TrimEnd('.', ' ', '\r', '\n');
                Console.WriteLine("Job Number:" + jobNumber);
                foundJobNumber = true;
            }

            // 2. get log file name
            // define the regular expression pattern
            string patternForLogFileName = @"^(.*\bMessages being written to\b.*)$";

            // use Regex.Match to find the match in the text
            Match matchForLogFileName = Regex.Match(disburseFeeJobLogMessage, patternForLogFileName, RegexOptions.Multiline);
            Console.WriteLine("matchForLogFileName.Success:" + matchForLogFileName.Success);

            // check if a match was found
            if (matchForLogFileName.Success)
            {
                // extract the matched line
                string matchedLine = matchForLogFileName.Groups[1].Value;
                Console.WriteLine("matchedLine:" + matchedLine);
                string[] stringSegements = matchedLine.Split('\\');
                Console.WriteLine("stringSegements:" + stringSegements[stringSegements.Length - 1]);
                logFileName = stringSegements[stringSegements.Length - 1].TrimEnd('.', '\r', ' ');
                Console.WriteLine("logFileName:" + logFileName);
                foundLogFileName = true;
            }
            return (foundJobNumber, jobNumber, foundLogFileName, logFileName);
        }

        // check payements logs for error - Import Receipts
        public (bool, string, string) checkPaymentsLogFile(string logFileType, string logFileLocation)
        {
            debug(TRACE, "checkPaymentsLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int fileType = 0;
            int numFilesProcessed = 0;
            int numFilesInvalid = 0;
            int numFilesValid = 0;
            int numFilesValidWithAnomalies = 0;
            int numRecordsCreated = 0;
            int numRecordsUpdated = 0;
            int numDetailLinesRead = 0;
            int numDetailLinesWithAnomalies = 0;
            int numLineInFile = 0;
            int numRecordsToBeCreated = 0;
            int numRecordsToBeUpdated = 0;

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // get fille name
            string[] splitString = logFileLocation.Split('\\');
            string logFileName = splitString[splitString.Length - 1];

            // detemine what type of file is passed
            if (logFileName.Contains("Import_Receipt"))
            {
                typeOfFile = "Import Receipt Log File";
                fileType = 1;
            }
            else if (logFileName.Contains("APOST"))
            {
                typeOfFile = "AusPost Log File";
                fileType = 2;
            }
            else if (logFileName.Contains("NAB"))
            {
                typeOfFile = "NAB Log File";
                fileType = 2;
            }
            else if (logFileName.Contains("PTUTSSAUACC"))
            {
                typeOfFile = "FlyWire WU Log File";
                fileType = 2;
            }
            else if (logFileName.Contains("PTUTSIACC"))
            {
                typeOfFile = "FlyWire CUST Log File";
                fileType = 2;
            }

            // read file and extract data
            foreach (string line in inputLines)
            {
                switch (fileType)
                {
                    //Import
                    case 1:
                        if (line.StartsWith("Number of files processed."))
                        {
                            numFilesProcessed = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[4]);
                        }
                        else if (line.StartsWith("Number of files invalid."))
                        {
                            numFilesInvalid = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[4]);
                        }
                        else if (line.StartsWith("Number of files valid."))
                        {
                            numFilesValid = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[4]);
                        }
                        else if (line.StartsWith("Number of files valid with anomalies."))
                        {
                            numFilesValidWithAnomalies = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[6]);
                        }
                        else if (line.StartsWith("Number of detail lines read."))
                        {
                            numDetailLinesRead = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[5]);
                        }
                        else if (line.StartsWith("Number of detail lines with anomalies."))
                        {
                            numDetailLinesWithAnomalies = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[6]);
                        }
                        else if (line.StartsWith("Number of records created."))
                        {
                            numRecordsCreated = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[4]);
                        }
                        else if (line.StartsWith("Number of records updated."))
                        {
                            numRecordsUpdated = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[4]);
                        }
                        break;

                    //AusPost, NAB, FlywWire WU, FlyWire CUST
                    case 2:
                        if (line.StartsWith("Number of lines in file."))
                        {
                            numLineInFile = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[5]);
                        }
                        else if (line.StartsWith("Number of detail lines read."))
                        {
                            numDetailLinesRead = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[5]);
                        }
                        else if (line.StartsWith("Number of detail lines with anomalies."))
                        {
                            numDetailLinesWithAnomalies = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[6]);
                        }
                        else if (line.StartsWith("Number of records to be created."))
                        {
                            numRecordsToBeCreated = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[6]);
                        }
                        else if (line.StartsWith("Number of records to be updated."))
                        {
                            numRecordsToBeUpdated = int.Parse(line.Split(new string[] { " " }, StringSplitOptions.None)[6]);
                        }
                        break;
                }
            }

            // error validation
            switch (fileType)
            {
                //Import
                case 1:
                    if (numFilesProcessed > 0)
                    {
                        // error #1
                        if (numFilesProcessed != numFilesValid)
                        {
                            message = "Test failed because:" +
                                "\nNumber of files processed: " + numFilesProcessed +
                                "\nNumber of files valid: " + numFilesValid +
                                "\nNumber of files invalid: " + numFilesInvalid;
                            result = false;
                            return (result, typeOfFile, message);
                        }
                        // error #2
                        if (numFilesValidWithAnomalies > 0)
                        {
                            message = "Test failed because:" +
                                "\nNumber of files processed: " + numFilesProcessed +
                                "\nNumber of files valid: " + numFilesValid +
                                "\nNumber of files valid with anomalies: " + numFilesValidWithAnomalies;
                            result = false;
                            return (result, typeOfFile, message);
                        }
                    }
                    else
                    {
                        message = "Test failed because:" + "\nNumber of files processed: 0";
                        result = false;
                        return (result, typeOfFile, message);
                    }
                    break;

                //AusPost, NAB, FlywWire WU, FlyWire CUST
                case 2:
                    // error #1
                    if (numDetailLinesRead != numLineInFile - 2)
                    {
                        message = "Test failed because:" +
                            "\nNumber of lines in file:" + numLineInFile +
                            "\nNumber of files invalid: " + numDetailLinesRead;
                        result = false;
                        return (result, typeOfFile, message);
                    }
                    // error #2
                    if (numDetailLinesWithAnomalies > 0)
                    {
                        message = "Test failed because:" +
                            "\nNumber of detail lines read: " + numDetailLinesRead +
                            "\nNumber of detail lines with anomalies: " + numDetailLinesWithAnomalies;
                        result = false;
                        return (result, typeOfFile, message);
                    }
                    break;
            }
            return (result, typeOfFile, message);
        }

        public (bool, string, string) checkStudyPackageAvailabilityRollForwardLogFile(string logFileLocation, bool wasTestrun)
        {
            debug(TRACE, "checkStudyPackageAvailabilityRollForwardLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int numAvailabilitiesSelected = 0;
            int numAvailabilitiesSuccess = 0;
            int numAvailabilitiesPartial = 0;
            int numAvailabilitiesNotSuccess = 0;
            string[] tempSplit = null;
            string tempReplace = "";

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.Contains("Availabilities Selected"))
                {
                    tempReplace = line.Replace(".", "");
                    tempSplit = tempReplace.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numAvailabilitiesSelected = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numAvailabilitiesSelected = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.Contains("Availabilities Successfully Rolled Forward"))
                {
                    tempReplace = line.Replace(".", "");
                    tempSplit = tempReplace.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numAvailabilitiesSuccess = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numAvailabilitiesSuccess = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.Contains("Availabilities Partially Successfully Rolled Forward"))
                {
                    tempReplace = line.Replace(".", "");
                    tempSplit = tempReplace.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numAvailabilitiesPartial = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numAvailabilitiesPartial = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.Contains("Availabilities Not Successfully Rolled Forward"))
                {
                    tempReplace = line.Replace(".", "");
                    tempSplit = tempReplace.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numAvailabilitiesNotSuccess = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numAvailabilitiesNotSuccess = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
            }

            // error validation
            if (numAvailabilitiesSuccess == 0) // need an error file to check how the errors look like... 
            {
                message = "Test failed :" +
                    "\nNumber of Availabilities Selected: " + numAvailabilitiesSelected +
                   "\nNumber of Availabilities Successfully Rolled Forward: " + numAvailabilitiesSuccess +
                   "\nNumber of Availabilities Partially Successfully Rolled Forward: " + numAvailabilitiesPartial +
                    "\nNumber of Availabilities Not Successfully Rolled Forward: " + numAvailabilitiesNotSuccess;
                result = false;
                return (result, typeOfFile, message);

            }
            else
            {
                message = "Test Passed :" +
                   "\nNumber of Availabilities Selected: " + numAvailabilitiesSelected +
                   "\nNumber of Availabilities Successfully Rolled Forward: " + numAvailabilitiesSuccess +
                   "\nNumber of Availabilities Partially Successfully Rolled Forward: " + numAvailabilitiesPartial +
                    "\nNumber of Availabilities Not Successfully Rolled Forward: " + numAvailabilitiesNotSuccess;
                result = true;
                return (result, typeOfFile, message);
            }

        }

        public (bool, string, string) checkBulkSanctionProcessingLog(string logFileLocation, bool wasTestrun)
        {
            debug(TRACE, "checkCommentImportLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int numTotalRecords = 0;
            int numTotalSuccessful = 0;
            int numTotalWarnings = 0;
            int numTotalErrors = 0;
            string[] tempSplit = null;

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Total Number of Records"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalRecords = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalRecords = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Records with Errors"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalErrors = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalErrors = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Records with Warnings"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalWarnings = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalWarnings = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Records Successful"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numTotalSuccessful = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                    else
                    {
                        numTotalSuccessful = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
            }

            // error validation
            if (numTotalErrors >= 1) // need an error file to check how the errors look like... 
            {
                message = "Test failed :" +
                    "\nNumber of records processed: " + numTotalRecords +
                    "\nNumber of records errors: " + numTotalErrors +
                    "\nNumber of records warnings: " + numTotalWarnings +
                    "\nNumber of records successful: " + numTotalSuccessful;
                result = false;
                return (result, typeOfFile, message);

            }
            else
            {
                message = "Test Passed :" +
                   "\nNumber of records processed: " + numTotalRecords +
                   "\nNumber of records errors: " + numTotalErrors +
                   "\nNumber of records warnings: " + numTotalWarnings +
                   "\nNumber of records successful: " + numTotalSuccessful;
                result = true;
                return (result, typeOfFile, message);
            }

        }

        public (bool, string, string) checkImportRewardCandidatesLogFile(string logFileLocation, bool wasTestrun)
        {
            debug(TRACE, "checkTACLogFile: start");

            bool result = true;
            string message = "Success";
            string typeOfFile = "";

            int numLinesFiles = 0;
            int numDetailLinesRead = 0;
            int numDetailLinesAnomalies = 0;
            int numRecordsCreated = 0;
            int numRecordsUpdated = 0;
            string[] tempSplit = null;

            string[] inputLines = File.ReadAllLines(logFileLocation);

            // read file and extract data
            foreach (string line in inputLines)
            {
                if (line.StartsWith("Number of lines in file"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numLinesFiles = int.Parse(tempSplit[tempSplit.Length - 2]);
                    }
                    else
                    {
                        numLinesFiles = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Number of detail lines read"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numDetailLinesRead = int.Parse(tempSplit[tempSplit.Length - 2]);
                    }
                    else
                    {
                        numDetailLinesRead = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Number of detail lines with anomalies"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numDetailLinesAnomalies = int.Parse(tempSplit[tempSplit.Length - 2]);
                    }
                    else
                    {
                        numDetailLinesAnomalies = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Number of records to be created"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numRecordsCreated = int.Parse(tempSplit[tempSplit.Length - 3]);
                    }
                    else
                    {
                        numRecordsCreated = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
                else if (line.StartsWith("Number of records to be updated"))
                {
                    tempSplit = line.Split(new string[] { " " }, StringSplitOptions.None);
                    if (wasTestrun)
                    {
                        numRecordsUpdated = int.Parse(tempSplit[tempSplit.Length - 3]);
                    }
                    else
                    {
                        numRecordsUpdated = int.Parse(tempSplit[tempSplit.Length - 1]);
                    }
                }
            }

            // error validation
            if (numRecordsCreated > numDetailLinesRead) // need an error file to check how the errors look like... 
            {
                message = "Test failed :" +
                    "\nNumber of records created: " + numRecordsCreated +
                   "\nNumber of files updated: " + numRecordsUpdated +
                    "\nNumber of files anomalies: " + numDetailLinesAnomalies;
                result = false;
                return (result, typeOfFile, message);

            }
            else
            {
                message = "Test Passed :" +
                       "\nNumber of records created: " + numRecordsCreated +
                       "\nNumber of files updated: " + numRecordsUpdated +
                       "\nNumber of files anomalies: " + numDetailLinesAnomalies;
                result = true;
                return (result, typeOfFile, message);
            }

        }
    }
}