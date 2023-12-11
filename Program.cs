using System;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Runtime.Remoting;
using ExcelDna.Integration;

namespace TestExcelIntegration
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            InitializeExcelIntegration();

            var xlRef = new ExcelReference(1, 1, 1, 1, IntPtr.Zero);
            Console.WriteLine(xlRef);
            Console.ReadLine();
        }

        static void InitializeExcelIntegration()
        {
            // Build a proxy object for the internal interface IIntegrationHost
            var internalType = typeof(ExcelReference).Assembly.GetType("ExcelDna.Integration.IIntegrationHost");
            var proxyIntegrationHost = new InterfaceImplementer(internalType, (MethodInfo info) =>
            {
                // Implement logic for when an IIntegrationHost method is called.
                Console.WriteLine($"{info.Name} method called in IIntegrationHost");

                // We only 'implement' the TryExcelImpl method, and return 'failed' in that case.
                if (info.Name == "TryExcelImpl")
                    return XlCall.XlReturn.XlReturnFailed;

                return null;
            }).GetTransparentProxy();

            // Call the internal static method ExcelIntegration.ConfigureHost via Reflection 
            var configureHost = typeof(ExcelIntegration).GetMethod("ConfigureHost", BindingFlags.Static | BindingFlags.NonPublic);
            configureHost.Invoke(null, new[] { proxyIntegrationHost });
        }
    }

    public class InterfaceImplementer : RealProxy, IRemotingTypeInfo
    {
        readonly Type _type;
        readonly Func<MethodInfo, object> _callback;

        public InterfaceImplementer(Type type, Func<MethodInfo, object> callback) : base(type)
        {
            _callback = callback;
            _type = type;
        }

        public override IMessage Invoke(IMessage msg)
        {
            var call = msg as IMethodCallMessage;

            if (call == null)
                throw new NotSupportedException();

            var method = (MethodInfo)call.MethodBase;

            return new ReturnMessage(_callback(method), null, 0, call.LogicalCallContext, call);
        }

        public bool CanCastTo(Type fromType, object o) => fromType == _type;

        public string TypeName { get; set; }
    }

    //class IntegrationHost // : IIntegrationHost
    //{
    //    public XlCall.XlReturn TryExcelImpl(int xlFunction, out object result, params object[] parameters)
    //    {
    //        result = null;
    //        return XlCall.XlReturn.XlReturnFailed;
    //    }
    //    public byte[] GetResourceBytes(string resourceName, int type) => null;
    //    public Assembly LoadFromAssemblyPath(string assemblyPath) => null;
    //    public Assembly LoadFromAssemblyBytes(byte[] assemblyBytes, byte[] pdbBytes) => null;
    //    public void RegisterMethods(List<MethodInfo> methods) { }
    //    public void RegisterMethodsWithAttributes(List<MethodInfo> methods, List<object> functionAttributes, List<List<object>> argumentAttributes) { }
    //    public void RegisterDelegatesWithAttributes(List<Delegate> delegates, List<object> functionAttributes, List<List<object>> argumentAttributes) { }
    //    public void RegisterLambdaExpressionsWithAttributes(List<LambdaExpression> lambdaExpressions, List<object> functionAttributes, List<List<object>> argumentAttributes) { }
    //    public void RegisterRtdWrapper(string progId, object rtdWrapperOptions, object functionAttribute, List<object> argumentAttributes) {}
    //    public int LPenHelper(int wCode, ref XlCall.FmlaInfo fmlaInfo) => -1;
    //}
}
