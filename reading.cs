using System;
using Slb.Ocean.Core;
using Slb.Ocean.Petrel;
using Slb.Ocean.Petrel.UI;
using Slb.Ocean.Petrel.Workflow;

namespace OceanReadingData
{
    /// <summary>
    /// This class will control the lifecycle of the Module.
    /// The order of the methods are the same as the calling order.
    /// </summary>
    public class reading : IModule
    {
        private Process m_readInstance;
        public reading()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        #region IModule Members

        /// <summary>
        /// This method runs once in the Module life; when it loaded into the petrel.
        /// This method called first.
        /// </summary>
        public void Initialize()
        {
            // TODO:  Add reading.Initialize implementation
        }

        /// <summary>
        /// This method runs once in the Module life. 
        /// In this method, you can do registrations of the not UI related components.
        /// (eg: datasource, plugin)
        /// </summary>
        public void Integrate()
        {
            // Register CommandHandler
            PetrelSystem.CommandManager.CreateCommand(OceanReadingData.CommandHandler.ID, new OceanReadingData.CommandHandler());
            // Register UIOEbar
            PetrelSystem.CommandManager.CreateCommand(OceanReadingData.UIOEbar.ID, new OceanReadingData.UIOEbar());
            // Register acc
            PetrelSystem.CommandManager.CreateCommand(OceanReadingData.acc.ID, new OceanReadingData.acc());
            // Register openexcel
            PetrelSystem.CommandManager.CreateCommand(OceanReadingData.openexcel.ID, new OceanReadingData.openexcel());
            // Register OceanReadingData.read
            OceanReadingData.read readInstance = new OceanReadingData.read();
            PetrelSystem.WorkflowEditor.Add(readInstance);
            m_readInstance = new Slb.Ocean.Petrel.Workflow.WorkstepProcessWrapper(readInstance);
            PetrelSystem.ProcessDiagram.Add(m_readInstance, "Plug-ins");

            // TODO:  Add reading.Integrate implementation
        }

        /// <summary>
        /// This method runs once in the Module life. 
        /// In this method, you can do registrations of the UI related components.
        /// (eg: settingspages, treeextensions)
        /// </summary>
        public void IntegratePresentation()
        {
            // Add Ribbon Configuration file
            PetrelSystem.ConfigurationService.AddConfiguration(OceanReadingData.Properties.Resources.OceanRibbonConfiguration);

            // TODO:  Add reading.IntegratePresentation implementation
        }

        /// <summary>
        /// This method runs once in the Module life.
        /// right before the module is unloaded. 
        /// It usually happens when the application is closing.
        /// </summary>
        public void Disintegrate()
        {
            PetrelSystem.ProcessDiagram.Remove(m_readInstance);
            // TODO:  Add reading.Disintegrate implementation
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            // TODO:  Add reading.Dispose implementation
        }

        #endregion

    }


}