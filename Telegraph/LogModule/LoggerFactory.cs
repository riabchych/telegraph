﻿using System;
using Telegraph.LogModule.Loggers;

namespace Telegraph.LogModule
{
    public class LoggerFactory
    {
        protected LoggerFactory()
        {
        }

        public static T Create<T>() where T : ILogger, new()
        {
            return Activator.CreateInstance<T>();
        }
    }
}