// ----------------------------------------------------------------------------
// Dataverse Logic Dependency Explorer  v1.2026.11.11
// Maintainer: Gagan Deep Singh  |  LinkedIn: https://www.linkedin.com/in/glowofsky/
// Host: XrmToolBox  |  Data platform: Microsoft Dataverse  |  Export format: Excel workbook (.xlsx)
// ----------------------------------------------------------------------------
// LOCAL DEVELOPMENT SHIM - NOT PART OF THE DELIVERABLE PLUGIN.
// A tiny stand-in for the MSTest attributes/asserts and a reflection runner so the unit tests can be executed with Mono in
// an environment that cannot restore NuGet packages. The real test project uses Microsoft MSTest.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace Microsoft.VisualStudio.TestTools.UnitTesting
{
    [AttributeUsage(AttributeTargets.Class)] public class TestClassAttribute : Attribute { }
    [AttributeUsage(AttributeTargets.Method)] public class TestMethodAttribute : Attribute { }
    [AttributeUsage(AttributeTargets.Method)] public class TestInitializeAttribute : Attribute { }
    [AttributeUsage(AttributeTargets.Method)] public class TestCleanupAttribute : Attribute { }

    public class AssertFailedException : Exception { public AssertFailedException(string m) : base(m) { } }

    public static class Assert
    {
        public static void IsTrue(bool c, string m = null) { if (!c) throw new AssertFailedException("IsTrue failed. " + m); }
        public static void IsFalse(bool c, string m = null) { if (c) throw new AssertFailedException("IsFalse failed. " + m); }
        public static void IsNull(object o, string m = null) { if (o != null) throw new AssertFailedException("IsNull failed: <" + o + ">. " + m); }
        public static void IsNotNull(object o, string m = null) { if (o == null) throw new AssertFailedException("IsNotNull failed. " + m); }
        public static void Fail(string m = null) { throw new AssertFailedException("Fail. " + m); }
        public static void AreEqual<T>(T e, T a, string m = null)
        {
            if (!object.Equals(e, a)) throw new AssertFailedException("AreEqual failed. Expected <" + e + "> Actual <" + a + ">. " + m);
        }
        public static void AreNotEqual<T>(T e, T a, string m = null)
        {
            if (object.Equals(e, a)) throw new AssertFailedException("AreNotEqual failed. Value <" + a + ">. " + m);
        }
        public static T ThrowsException<T>(Action a) where T : Exception
        {
            try { a(); } catch (T ex) { return ex; } catch (Exception ex) { throw new AssertFailedException("Expected " + typeof(T).Name + " but got " + ex.GetType().Name); }
            throw new AssertFailedException("Expected exception " + typeof(T).Name + " was not thrown.");
        }
    }

    public static class StringAssert
    {
        public static void Contains(string v, string s, string m = null)
        {
            if (v == null || v.IndexOf(s, StringComparison.Ordinal) < 0) throw new AssertFailedException("Expected <" + v + "> to contain <" + s + ">. " + m);
        }
        public static void DoesNotMatch(string v, System.Text.RegularExpressions.Regex r) { if (r.IsMatch(v)) throw new AssertFailedException("Unexpected match"); }
    }

    public static class MiniRunner
    {
        public static int Run(Assembly asm)
        {
            int passed = 0, failed = 0;
            var failures = new List<string>();
            foreach (Type t in asm.GetTypes().Where(x => x.GetCustomAttribute<TestClassAttribute>() != null).OrderBy(x => x.Name))
            {
                foreach (MethodInfo m in t.GetMethods().Where(x => x.GetCustomAttribute<TestMethodAttribute>() != null).OrderBy(x => x.Name))
                {
                    object inst = Activator.CreateInstance(t);
                    MethodInfo init = t.GetMethods().FirstOrDefault(x => x.GetCustomAttribute<TestInitializeAttribute>() != null);
                    MethodInfo clean = t.GetMethods().FirstOrDefault(x => x.GetCustomAttribute<TestCleanupAttribute>() != null);
                    var sw = Stopwatch.StartNew();
                    try
                    {
                        if (init != null) init.Invoke(inst, null);
                        m.Invoke(inst, null);
                        passed++;
                        Console.WriteLine("  PASS  " + t.Name + "." + m.Name);
                    }
                    catch (TargetInvocationException tie)
                    {
                        failed++;
                        string msg = t.Name + "." + m.Name + ": " + tie.InnerException.GetType().Name + ": " + tie.InnerException.Message;
                        failures.Add(msg + "\n" + tie.InnerException.StackTrace);
                        Console.WriteLine("  FAIL  " + msg);
                    }
                    finally { if (clean != null) { try { clean.Invoke(inst, null); } catch { } } }
                }
            }
            Console.WriteLine();
            Console.WriteLine("Passed: " + passed + "  Failed: " + failed + "  Total: " + (passed + failed));
            foreach (string f in failures) Console.WriteLine("\n--- " + f);
            return failed == 0 ? 0 : 1;
        }

        public static int Main(string[] args) { return Run(Assembly.GetEntryAssembly()); }
    }
}
