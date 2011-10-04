﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SwDocumentMgr;

namespace swLastSaved
{
    class Program
    {
        static void Main(string[] args)
        {

            string docPath;
            bool quietMode;
            switch (args.Length)
            {
                case 1:
                    quietMode = false;
                    docPath = args[0];
                    break;
                case 2:
                    if (args[0] != "/q")
                    {
                        quietMode = false;
                        inputError(quietMode);
                        return;
                    }
                    quietMode = true;
                    docPath = args[1];
                    break;
                default:
                    quietMode = false;
                    inputError(quietMode);
                    return;
            }
            //Get Document Type
            SwDmDocumentType docType = setDocType(docPath);

            SwDMClassFactory dmClassFact = new SwDMClassFactory();
            SwDMApplication dmDocManager = dmClassFact.GetApplication("<Insert your key here.>") as SwDMApplication;
            SwDmDocumentOpenError OpenError;
            SwDMDocument dmDoc = dmDocManager.GetDocument(docPath, docType, true, out OpenError) as SwDMDocument;


            string LastSavedByUser;

            if (dmDoc != null)
            {
                try
                {
                    LastSavedByUser = dmDoc.LastSavedBy;
                    if (LastSavedByUser == null)
                    {
                        Console.WriteLine(docPath + "\t " + LastSavedByUser + ".  This probably isn't a SolidWorks file or it is severely damaged.");
                        return;
                    }
                    Console.WriteLine(docPath + "\t" + LastSavedByUser);
                }
                catch (Exception e)
                {
                    Console.WriteLine(docPath + "\t" + "Could not get file version. " + e.Message + e.StackTrace);
                    return;
                }
                dmDoc.CloseDoc();
            }
            else
            {
                switch (OpenError)
                {
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorFail:
                        Console.WriteLine(docPath + "\tFile failed to open; reasons could be related to permissions, the file is in use, or the file is corrupted.");
                        inputError(quietMode);
                        break;
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorFileNotFound:
                        Console.WriteLine(docPath + "\tFile not found");
                        inputError(quietMode);
                        break;
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorNonSW:
                        Console.WriteLine(docPath + "\tNon-SolidWorks file was opened");
                        inputError(quietMode);
                        break;
                    default:
                        Console.WriteLine(docPath + "\tAn unknown errror occurred.  Something is wrong with the code of \"GetSolidWorksFileVersion\"");
                        inputError(quietMode);
                        break;
                }
            }




        }



        static SwDmDocumentType setDocType(string docPath)
        {
            string fileExtension = System.IO.Path.GetExtension(docPath);

            //Notice no break statement is needed because I used return to get out of the switch statement.
            switch (fileExtension.ToUpper())
            {
                case ".SLDASM":
                case ".ASM":
                    return SwDmDocumentType.swDmDocumentAssembly;
                case ".SLDDRW":
                case ".DRW":
                    return SwDmDocumentType.swDmDocumentDrawing;
                case ".SLDPRT":
                case ".PRT":
                    return SwDmDocumentType.swDmDocumentPart;
                default:
                    return SwDmDocumentType.swDmDocumentUnknown;
            }

        }



        static void inputError(bool quietMode)
        {
            if (quietMode)
                return;

            Console.WriteLine(@"

Returns the user that last saved the document.

Syntax 
    [option] [SolidWorksFilePath]

Output
    SolidWorksFilePath  LastSavedByUser

Options
    /q      Quiet mode.  Suppresses the current message.  It does
            not suppress the one line error messages related to problems
            opening SolidWorks Files.  Quiet mode is useful for batch files
            when you are directing the output to a file.  The main error 
            message is suppressed but you are still informed about problems 
            opening files.

Version 2011-Oct-04 18:24
Written and Maintained by Jason Nicholson
http://github.com/jasonnicholson/swLastSaved");
        }


    }
}
