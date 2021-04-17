using System;
using System.Collections.Generic;
using MGCPCB;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;

namespace Mentor_MountingHole_Shaver
{

    class Program
    {
        static MGCPCB.Document pcbDoc;
        static PadstackEditorLib.PadstackEditorDlg _dlg;

        [STAThread]
        static void Main(string[] args)
        {
            string _appGUID = null;
            if (args.Length == 2)
            {
                if (args[0] == "-guid" && args[1] != null)
                {
                    _appGUID = args[1];
                }
            }

            #region Instance Connection Code
            try
            {
                MGCPCBReleaseEnvironmentLib.IMGCPCBReleaseEnvServer _server =
                    (MGCPCBReleaseEnvironmentLib.IMGCPCBReleaseEnvServer)Activator.CreateInstance(
                        Marshal.GetTypeFromCLSID(
                            new Guid("44983CB8-19B0-4695-937A-6FF0B74ECFC5")
                        )
                    );


                _server.SetEnvironment("");
                string VxVersion = _server.sddVersion;
                string strSDD_HOME = _server.sddHome;
                int length = strSDD_HOME.IndexOf("SDD_HOME");
                strSDD_HOME = strSDD_HOME.Substring(0, length).Replace("\\", "\\\\") + "SDD_HOME";
                _server.SetEnvironment(strSDD_HOME);
                string progID = _server.ProgIDVersion;

                object[,] _releases = (object[,])_server.GetInstalledReleases();
                dynamic pcbApp = null;

                for (int i = 1; i < _releases.Length / 4; i++)
                {
                    string _com_version = Convert.ToString(_releases[i, 0]);
                    try
                    {
                        pcbApp = Interaction.GetObject(null, "MGCPCB.Application." + _com_version);

                        pcbDoc = pcbApp.ActiveDocument;
                        dynamic licApp = Interaction.CreateObject("MGCPCBAutomationLicensing.Application." + _com_version);
                        int _token = licApp.GetToken(pcbDoc.Validate(0));
                        pcbDoc.Validate(_token);

                        break;
                    }
                    catch (Exception m)
                    {
                    }
                }


                if (pcbApp == null)
                {
                    System.Windows.Forms.MessageBox.Show("Could not found active Xpedition or PADSPro Application");
                    System.Environment.Exit(1);
                }



            }
            catch (Exception m)
            {
                System.Windows.Forms.MessageBox.Show(m.Message + "\r\n" + m.Source + "\r\n" + m.StackTrace);
            }
            #endregion
#if !DEBUG
            try
            {
                pcbDoc.TransactionStart(EPcbDRCMode.epcbDRCModeNone);
#endif
            _dlg = (PadstackEditorLib.PadstackEditorDlg)pcbDoc.PadstackEditor;
            //_dlg.LockServer();
            MountingHoles _holes = pcbDoc.get_MountingHoles(EPcbSelectionType.epcbSelectSelected);
            int _count = 0;
            pcbDoc.Application.Gui.ProgressBarInitialize(true, "Shaving Pads", _holes.Count, 0);

            List<object[,]> _inShape = new List<object[,]>();
            _inShape.Add((object[,])pcbDoc.RouteBorder.Geometry.get_PointsArray(EPcbUnit.epcbUnitMils));

            List<object[,]> _outShape = new List<object[,]>();
            foreach (Contour _cont in pcbDoc.get_Contours(EPcbSelectionType.epcbSelectAll))
            {
                try
                {
                    _outShape.Add(xFS_DrawingTools.DrawingTools.V2_Oversize((object[,])_cont.Geometry.get_PointsArray(EPcbUnit.epcbUnitMils), 13));
                }
                catch
                {

                }
            }



            foreach (Obstruct _ob in pcbDoc.get_Obstructs(EPcbObstructType.epcbObstructTraceVia, EPcbSelectionType.epcbSelectAll))
                _outShape.Add((object[,])_ob.Geometry.get_PointsArray(EPcbUnit.epcbUnitMils));

            foreach (MountingHole _hole in _holes)
            {
                Net _n = _hole.Net;
                pcbDoc.Application.Gui.ProgressBar(_count);
                #region Working with Padstacks
                PadstackEditorLib.Pad _newPad = _dlg.ActiveDatabase.NewPad();
                var _pad = _hole.get_Pads()[1];
                Geometry _geom = _pad.Geometries[1];

                var _padList = xFS_DrawingTools.DrawingTools.V2_CutShape((object[,])_geom.get_PointsArray(EPcbUnit.epcbUnitMils), _inShape, _outShape, 0, 0);
                var _parr = xFS_DrawingTools.DrawingTools.V2_Transform(_padList[0], -_hole.get_PositionX(EPcbUnit.epcbUnitMils), -_hole.get_PositionY(EPcbUnit.epcbUnitMils));
                _newPad.Shape = PadstackEditorLib.EPsDBPadShape.epsdbPadShapeCustom;
                _newPad.PutGeometry(_parr.Length / 3, _parr, PadstackEditorLib.EPsDBUnit.epsdbUnitMils);
                _newPad.Name = FindNameForPad(pcbDoc, _dlg.ActiveDatabase.get_Pads(), "Custom_" + _hole.CurrentPadstack.Name);
                _newPad.Commit();

                PadstackEditorLib.Pad _smPad = _dlg.ActiveDatabase.NewPad();
                Geometry _smgeom = _pad.Geometries[1];
                var _smList = xFS_DrawingTools.DrawingTools.V2_CutShape((object[,])_smgeom.get_PointsArray(EPcbUnit.epcbUnitMils), _inShape, _outShape, 0, 0);
                var _smArr = xFS_DrawingTools.DrawingTools.V2_Transform(_smList[0], -_hole.get_PositionX(EPcbUnit.epcbUnitMils), -_hole.get_PositionY(EPcbUnit.epcbUnitMils));
                _smPad.Shape = PadstackEditorLib.EPsDBPadShape.epsdbPadShapeCustom;
                _smPad.PutGeometry(_smArr.Length / 3, _smArr, PadstackEditorLib.EPsDBUnit.epsdbUnitMils);
                _smPad.Name = FindNameForPad(pcbDoc, _dlg.ActiveDatabase.get_Pads(), "Custom_SM_" + _hole.CurrentPadstack.Name);
                _smPad.Commit();

                PadstackEditorLib.Padstack _ps = _dlg.ActiveDatabase.FindPadstack(_hole.CurrentPadstack.Name).Copy();
                _ps.set_Pad(PadstackEditorLib.EPsDBPadLayer.epsdbPadLayerMountSide, _newPad);
                _ps.set_Pad(PadstackEditorLib.EPsDBPadLayer.epsdbPadLayerOppositeSide, _newPad);
                _ps.set_Pad(PadstackEditorLib.EPsDBPadLayer.epsdbPadLayerTopMountSoldermask, _smPad);
                _ps.set_Pad(PadstackEditorLib.EPsDBPadLayer.epsdbPadLayerBottomMountSoldermask, _smPad);

                if (pcbDoc.LayerCount > 2)
                {
                    PadstackEditorLib.Pad _innerPad = _dlg.ActiveDatabase.NewPad();
                    var _innerpad = _hole.get_Pads(2)[1];

                    if (_innerpad != null)
                    {
                        Geometry _innergeom = _innerpad.Geometries[1];
                        var _innerList = xFS_DrawingTools.DrawingTools.V2_CutShape((object[,])_innergeom.get_PointsArray(EPcbUnit.epcbUnitMils), _inShape, _outShape, 0, 0);
                        var _innerparr = xFS_DrawingTools.DrawingTools.V2_Transform(_innerList[0], -_hole.get_PositionX(EPcbUnit.epcbUnitMils), -_hole.get_PositionY(EPcbUnit.epcbUnitMils));
                        _innerPad.Shape = PadstackEditorLib.EPsDBPadShape.epsdbPadShapeCustom;
                        _innerPad.PutGeometry(_innerparr.Length / 3, _innerparr, PadstackEditorLib.EPsDBUnit.epsdbUnitMils);
                        _innerPad.Name = FindNameForPad(pcbDoc, _dlg.ActiveDatabase.get_Pads(), "Custom_INNER_" + _hole.CurrentPadstack.Name);
                        _innerPad.Commit();
                        _ps.set_Pad(PadstackEditorLib.EPsDBPadLayer.epsdbPadLayerInternal, _innerPad);
                    }
                }

                _ps.Name = FindNameForPadstack(pcbDoc, "Modified_" + _hole.CurrentPadstack.Name);

                double _X = _hole.get_PositionX(EPcbUnit.epcbUnitMils);
                double _Y = _hole.get_PositionY(EPcbUnit.epcbUnitMils);
                _dlg.SaveActiveDatabase();
                //_dlg.UnlockServer();
                _hole.Delete();
                Padstack _p = pcbDoc.PutPadstack(1, pcbDoc.LayerCount, _ps.Name, false, false);
                pcbDoc.PutMountingHole(_X, _Y, _p, null, null, EPcbAnchorType.epcbAnchorNone, EPcbUnit.epcbUnitMils).Net = _n;
                #endregion
                _count++;
            }

            pcbDoc.Application.Gui.ProgressBarInitialize(false);
#if !DEBUG 
            }
            catch ( Exception m )
            {
                System.Windows.Forms.MessageBox.Show(m.Message + "\r\n" + m.Source);
                pcbDoc.Application.Gui.StatusBarText(m.Message + "\t" + m.Source, EPcbStatusField.epcbStatusFieldError);
            }
            finally
            {
     
                if ( _dlg != null )
                {
                    //_dlg.UnlockServer();
                    _dlg.Quit();
                }
                pcbDoc.TransactionEnd();
                //pcbDoc.Application.UnlockServer();
                pcbDoc.Application.Gui.ProgressBarInitialize(false);

            }
#endif
        }

        static string FindNameForPad(MGCPCB.Document pcbDoc, PadstackEditorLib.Pads _pads, string _padName)
        {
            List<string> _padNames = new List<string>();
            foreach (PadstackEditorLib.Pad _pad in _pads)
            {
                _padNames.Add(_pad.Name);
            }

            int counter = 0;
            string _newPadName = _padName;
            while (_padNames.Contains(_newPadName))
            {
                _newPadName = _padName + "_" + counter;
                counter++;
            }
            return _newPadName;

        }
        static string FindNameForPadstack(MGCPCB.Document pcbDoc, string _padName)
        {
            object[] _pads = (object[])pcbDoc.get_PadstackNames();

            List<string> _names = new List<string>();
            foreach (object _pad in _pads)
            {
                _names.Add(_pad.ToString());
            }
            int counter = 0;
            string _newPadName = _padName;
            while (_names.Contains(_newPadName))
            {
                _newPadName = _padName + "_" + counter;
                counter++;
            }

            return _newPadName;
        }
    }
}
