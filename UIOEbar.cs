using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Slb.Ocean.Petrel.Commands;
using Slb.Ocean.Petrel;

namespace OceanReadingData
{
    class UIOEbar : SimpleCommandHandler
    {
        public static string ID = "OceanReadingData.NewCommand3";

        #region SimpleCommandHandler Members

        public override bool CanExecute(Slb.Ocean.Petrel.Contexts.Context context)
        { 
            return true;
        }

        public override void Execute(Slb.Ocean.Petrel.Contexts.Context context)
        {          
            //TODO: Add command execution logic here
            PetrelLogger.InfoOutputWindow(string.Format("{0} clicked", @"OpenUI" ));
            Form1 fr = new Form1();
            PetrelSystem.ShowModeless(fr);
        }
    
        #endregion
    }
}
