// Guids.cs
// MUST match guids.h
using System;

namespace Firefly.Box.VSIntegration
{
    static class GuidList
    {
        public const string guidFirefly_Box_VSIntegrationPkgString = "bd955b23-3f46-494e-b2d7-266868d4cded";
        public const string guidFirefly_Box_VSIntegrationCmdSetString = "2763440e-4de7-40d3-95ed-4fbbf706848b";
        public const string guidToolWindowPersistanceString = "8f28e43a-e6c8-471b-8d1b-7af0ce85884f";
        public const string guidFirefly_Box_VSIntegrationEditorFactoryString = "2d5f8cf3-7e68-4fe1-a8ab-f1a90c7a7361";

        public static readonly Guid guidFirefly_Box_VSIntegrationPkg = new Guid(guidFirefly_Box_VSIntegrationPkgString);
        public static readonly Guid guidFirefly_Box_VSIntegrationCmdSet = new Guid(guidFirefly_Box_VSIntegrationCmdSetString);
        public static readonly Guid guidToolWindowPersistance = new Guid(guidToolWindowPersistanceString);
        public static readonly Guid guidFirefly_Box_VSIntegrationEditorFactory = new Guid(guidFirefly_Box_VSIntegrationEditorFactoryString);
    };
}