using System;
using System.Windows.Forms;

namespace Microsoft.Internal.MSContactImporter
{
    internal static class AbortRetryIgnorePattern
    {
        internal static void CallMethod(Action action, IState state, Action cancelAction, Func<Exception, IState, DialogResult> answer)
        {
            bool flag = true;
            while (flag)
            {
                try
                {
                    action();
                    flag = false;
                }
                catch (Exception arg)
                {
                    Logger.LogMessageToConsole(state.ToString(), arg);

                    switch (answer(arg, state))
                    {
                        case DialogResult.Abort:
                            flag = false;
                            cancelAction();
                            break;

                        case DialogResult.Retry:
                            flag = true;
                            break;

                        case DialogResult.Ignore:
                            flag = false;
                            break;
                    }
                }
            }
        }

        internal static T CallFunction<T>(Func<T> action, IState state, Action cancelAction, Func<Exception, IState, DialogResult> answer)
        {
            bool flag = true;
            T result = default(T);
            while (flag)
            {
                try
                {
                    result = action();
                    flag = false;
                }
                catch (Exception arg)
                {
                    switch (answer(arg, state))
                    {
                        case DialogResult.Abort:
                            flag = false;
                            cancelAction();
                            break;

                        case DialogResult.Retry:
                            flag = true;
                            break;

                        case DialogResult.Ignore:
                            flag = false;
                            break;
                    }
                }
            }
            return result;
        }
    }
}