namespace Microsoft.Internal.MSContactImporter
{
    internal interface IState
    {
    }

    internal class State : IState
    {
        public string Function
        {
            get;
            set;
        }
    }
}