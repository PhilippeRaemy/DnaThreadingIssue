namespace Cxl.GenericToolbox.DNA
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using Dna.ThreadingIssue;
    using MoreLinq;
    using static System.Diagnostics.Trace;

    static class Tracer
    {
        static readonly string FileName;

        static Tracer()
        {
            var logFolder = new DirectoryInfo(@"c:\data\log\ThreadingIssue");
            logFolder.Create();
            logFolder.GetFiles("*.log", SearchOption.TopDirectoryOnly)
                .Where(f => f.CreationTimeUtc < DateTime.UtcNow.AddDays(-15))
                .ToArray()
                .ForEach(f => f.Delete());
            FileName = Path.Combine(logFolder.FullName, $"ThreadingIssue{DateTime.UtcNow:yyyy-MM-dd_HH-mm-ss}.log");
            Listeners.Add(new TextWriterTraceListener(FileName, "ThreadingIssueListener"));
            AutoFlush = true;
            Info($"Initialize log {FileName}");
        }

        public static void DisplayTrace()
        {
            Process.Start(FileName);
        }

        static string CallingMethod(int depth)
        {
            var methodBase = new StackTrace().GetFrame(depth).GetMethod();
            return $"{methodBase.DeclaringType?.Name}.{methodBase.Name}";
        }

        static void Write(string severity, IEnumerable<object> message, int depth) 
            => WriteLine($"{DateTime.UtcNow:s} {severity} {CallingMethod(depth)} {message.ToDelimitedString(" ")}");

        public static void Info (params object[] message) => Write("Info ", message, 3);
        public static void Debug(params object[] message) => Write("Debug", message, 3);
        public static void Error(params object[] message) => Write("Error", message, 3);

        public static void Info (Action action, params object[] message) { Write("Info ", message, 3); action?.Invoke(); }
        public static void Debug(Action action, params object[] message) { Write("Debug", message, 3); action?.Invoke(); }
        public static void Error(Action action, params object[] message) { Write("Error", message, 3); action?.Invoke(); }

        public static T Info <T>(Func<T> func, params object[] message) { Write("Info ", message, 3); return func(); }
        public static T Debug<T>(Func<T> func, params object[] message) { Write("Debug", message, 3); return func(); }
        public static T Error<T>(Func<T> func, params object[] message) { Write("Error", message, 3); return func(); }

        // following methods are for generic wrappers (such as CatchException)
        public static void InfoDeep (params object[] message) => Write("Info ", message, 4);
        public static void DebugDeep(params object[] message) => Write("Debug", message, 4);
        public static void ErrorDeep(params object[] message) => Write("Error", message, 4);

        public static void InfoDeep (Action action, params object[] message) { Write("Info ", message, 4); action?.Invoke(); }
        public static void DebugDeep(Action action, params object[] message) { Write("Debug", message, 4); action?.Invoke(); }
        public static void ErrorDeep(Action action, params object[] message) { Write("Error", message, 4); action?.Invoke(); }

        public static T InfoDeep <T>(Func<T> func , params object[] message) { Write("Info ", message, 4); return func(); }
        public static T DebugDeep<T>(Func<T> func , params object[] message) { Write("Debug", message, 4); return func(); }
        public static T ErrorDeep<T>(Func<T> func , params object[] message) { Write("Error", message, 4); return func(); }
    }
}