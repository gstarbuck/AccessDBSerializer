using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AccessDBSerializer
{
    public enum AccessObjectType
    {
        acDefault = -1,
        acDiagram = 8,
        acForm = 2,
        acFunction = 10,
        acMacro = 4,
        acModule = 5,
        acQuery = 1,
        acReport = 3,
        acServerView = 7,
        acStoredProcedure = 9,
        acTable = 0,
        acCmdCompileAndSaveAllModules = 125
    }
}
