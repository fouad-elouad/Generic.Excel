﻿using Generic.Excel.Samples.Utilities;
using RandomSharp;
using RandomSharp.Impl;

namespace Generic.Excel.Samples
{
    public abstract class SampleBase
    {
        protected DataFactory _DataFactory;
        protected const string _ExcelExtension = ".xlsx";

        public SampleBase(DataFactory dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public SampleBase()
        {
            IRandomizer randomizer = new DefaultRandomizer();
            _DataFactory = new DataFactory(randomizer);
        }

        protected string FilePath(string fileName)
        {
            string directoryPath = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(directoryPath, fileName);
        }

        public abstract void Run();
    }
}
