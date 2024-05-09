using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class FileBlock
    {
        internal string fileName;
        internal List<Block> blockList;
    }

    internal class Block
    {
        internal List<string> rows;
    }
}
