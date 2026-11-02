using System;
internal sealed class ContextForegroundResult
{
    internal bool InitiallyForeground,RequestAccepted,Verified;
}
internal static class ContextForegroundPolicy
{
    internal static ContextForegroundResult Establish(Func<bool> isForeground,Func<bool> request,Action pump,Func<long> elapsed,long timeout)
    {
        if(timeout<1)throw new ArgumentOutOfRangeException("timeout");
        var result=new ContextForegroundResult {InitiallyForeground=isForeground()};
        if(result.InitiallyForeground){result.Verified=true;return result;}
        result.RequestAccepted=request();
        while(!isForeground()){if(elapsed()>=timeout)return result;pump();}
        result.Verified=true;return result;
    }
}