using System;

namespace InSolve.dmach.Mail
{
    public interface IAttachmentConstructor
    {
        System.IO.Stream Stream { get; }
        void CompleteConstruction();
    }
}
