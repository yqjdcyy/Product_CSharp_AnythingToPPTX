using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnythingToPPTX.Utils
{
    class ConsoleStdIO : Ghostscript.NET.GhostscriptStdIO
    {
        public ConsoleStdIO(bool handleStdIn, bool handleStdOut, bool handleStdErr) : base(handleStdIn, handleStdOut, handleStdErr) { }

        public override void StdIn(out string input, int count)
        {
            char[] userInput = new char[count];
            Console.In.ReadBlock(userInput, 0, count);
            input = new string(userInput);
        }

        public override void StdOut(string output)
        {
            Console.Write(output);
        }

        public override void StdError(string error)
        {
            Console.Write(error);
        }
    }
}
