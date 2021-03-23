#region Copyright © 2015 KOCH Supply and Trading SARL
// --------------------------------------------------
// Copyright © 2015 KOCH Supply and Trading SARL
//
// Warning: This source code is protected by copyright 
// law.  Unauthorized reproduction or distribution of 
// any part of this code may result in severe civil 
// and criminal penalties, and will be prosecuted to 
// the maximum extent possible under the law.
// --------------------------------------------------
#endregion

namespace Dna.ThreadingIssue
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Cxl.GenericToolbox.DNA;
    using ExcelDna.Integration;
    using Excel = NetOffice.ExcelApi;

    // ReSharper disable once ClassNeverInstantiated.Global : instantiated by DNA
    public class ThreadingIssue : IExcelAddIn
    {
        const bool USE_QUEUE_AS_MACRO = false;

        public static Excel.Workbook ThisWorkbook
        {
            get
            {
                try
                {
                    return ExcelApplication.ActiveWorkbook;
                }
                catch
                {
                    return null;
                }
            }
        }

        static Excel.Application _excelApplication;
        static Excel.Application ExcelApplication => _excelApplication ??= new Excel.Application(null, ExcelDnaUtil.Application);

        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => ex.ToString());

            ExcelApplication.WorkbookActivateEvent    += App_WorkbookActivateEvent;
            ExcelApplication.WorkbookDeactivateEvent  += App_WorkbookDeactivateEvent;
            ExcelApplication.WorkbookBeforeCloseEvent += App_WorkbookBeforeCloseEvent;
            ExcelApplication.WorkbookAfterSaveEvent   += App_WorkbookAfterSaveEvent;
            ExcelApplication.WorkbookBeforeSaveEvent  += App_WorkbookBeforeSaveEvent;
            ExcelApplication.WorkbookOpenEvent        += App_WorkbookOpenEvent;
            Tracer.Info();
        }

        public void AutoClose()
        {
            Tracer.Debug();
        }

        static void App_WorkbookOpenEvent(Excel.Workbook wb)
        {
            Tracer.Debug(wb.FullName); 
            App_WorkbookActivateEvent(wb);
        }

        static void App_WorkbookBeforeSaveEvent(Excel.Workbook wb, bool saveAsUi, ref bool cancel)
        {
            Tracer.Debug(wb.FullName, nameof(saveAsUi), saveAsUi);
            if (saveAsUi) App_WorkbookBeforeCloseEvent(wb, ref cancel);
        }

        static void App_WorkbookAfterSaveEvent(Excel.Workbook wb, bool success)
        {
            Tracer.Debug(wb.FullName);
            App_WorkbookActivateEvent(wb);
        }

        static void App_WorkbookBeforeCloseEvent(Excel.Workbook wb, ref bool cancel)
        {
            /* it's not the right place here to de-register the workbook because this event is fired before the user has a chance to 
             * cancel the close.
             * ==> Moving the logic to activate and deactivate events
             * https://social.msdn.microsoft.com/Forums/vstudio/en-US/39572297-6516-4a8a-9ca6-211a87e78d7b/get-workbook-close-event?forum=vsto
            */
            Tracer.Debug(wb.FullName);
        }


        static void App_WorkbookDeactivateEvent(Excel.Workbook wb)
        {
            Tracer.Debug();
            RunnerIsAlive = false;
        }

        static bool RunnerIsAlive;

        static void App_WorkbookActivateEvent(Excel.Workbook wb)
        {
            Tracer.Debug();
            RunnerIsAlive = true;
            if (USE_QUEUE_AS_MACRO)
                ExcelAsyncUtil.QueueAsMacro(() => MacroRunner(wb, wb.Name));
            else new Thread(() => ThreadRunner(wb, wb.Name)).Start();
        }

        static async void ThreadRunner(Excel.Workbook wb, string name)
        {
            Tracer.Debug($"Starting {name}");
            while (RunnerIsAlive)
            {
                try
                {
                    var _ = 1; //  wb.ActiveSheet;
                }
                catch (Exception) // if the workbook has no active sheet it means we've got an invalid workbook pointer, probably closed!
                {
                    Tracer.Debug($"Workbook has no active sheet: it means we've got an invalid workbook pointer, Probably closed!  {name}");
                    Tracer.Debug($"Quit {name}");
                    RunnerIsAlive = false;
                    break;
                }

                Tracer.Debug($"Running {name}");
                await Task.Delay(1000);
            }
            Tracer.Debug($"Done {name}");
        }

        static async void MacroRunner(Excel.Workbook wb, string name)
        {
            Tracer.Debug($"Starting {name}");
            if (!RunnerIsAlive)
            {
                Tracer.Debug($"Stopping {name}"); 
                return;
            }

            RunnerIsAlive = false; // notify add-in that there's no more queued macro for this runner

            try
            {
                var _ = 1; // wb.ActiveSheet;
            }
            catch (Exception) // if the workbook has no active sheet it means we've got an invalid workbook pointer, probably closed!
            {
                Tracer.Debug($"Workbook has no active sheet: it means we've got an invalid workbook pointer, Probably closed!  {name}");
                Tracer.Debug($"Quit {name}");
                return;
            }

            Tracer.Debug($"Running {name}");
            await Task.Delay(1000);
            RunnerIsAlive = true; // notify add-in that there's again a queued macro for this runner
            ExcelAsyncUtil.QueueAsMacro(() => MacroRunner(wb, name));
            Tracer.Debug($"Requeued {name}");
        }
    }
}
