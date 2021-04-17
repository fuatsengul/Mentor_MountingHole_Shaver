using System.Collections.Generic;
using System.Runtime.InteropServices;
using MASKENGINEAUTOMATIONCONTROLLERLib;

namespace xFS_DrawingTools
{
    public static class DrawingTools
    {

        public static object[,] V2_Oversize( object[,] input, double offsetAsMils, int oversizeType = 0 )
        {
            EMaskEngineOversizeType _type = EMaskEngineOversizeType.emeOversizeTypeRound;
            if(oversizeType!= 0)
                _type = EMaskEngineOversizeType.emeOversizeTypeSquare;
            MaskEngine _me = new MaskEngine();
            

            Mask _BAmask = _me.Masks.Add();
            _BAmask.Shapes.AddByPointsArray(input.Length / 3, input, true, EMaskEngineUnit.emeUnitMils);
            _BAmask.Oversize(offsetAsMils, _type, EMaskEngineUnit.emeUnitMils);

            object[,] retVal  = (object[,])_BAmask.Shapes.Item[1].get_PointsArray(MASKENGINEAUTOMATIONCONTROLLERLib.EMaskEngineUnit.emeUnitMils);
            Marshal.FinalReleaseComObject(_me);
            return retVal;
        }


        public static List<object[,]> V2_CutShape( object[,] input, List<object[,]> inShapes, List<object[,]> outShapes, double inShapeOffset, double outShapeOffset )
        {
            MaskEngine _me = new MaskEngine();
            Mask _inShapeMask = _me.Masks.Add();
            Mask _outShapeMask = _me.Masks.Add();



            if ( inShapes != null )
            {
                foreach ( object[,] _cutter in inShapes )
                {
                    _inShapeMask.Shapes.AddByPointsArray(_cutter.Length / 3, _cutter, true, EMaskEngineUnit.emeUnitMils);
                }
                if ( inShapeOffset != 0 )
                    _inShapeMask.Oversize(inShapeOffset, EMaskEngineOversizeType.emeOversizeTypeRound, EMaskEngineUnit.emeUnitMils);
            }
            if ( outShapes != null )
            {
                foreach ( object[,] _cutter in outShapes )
                {
                    _outShapeMask.Shapes.AddByPointsArray(_cutter.Length / 3, _cutter, true, EMaskEngineUnit.emeUnitMils);
                }
                if ( outShapeOffset != 0 )
                    _outShapeMask.Oversize(outShapeOffset, EMaskEngineOversizeType.emeOversizeTypeRound, EMaskEngineUnit.emeUnitMils);
            }


            Mask _objMask = _me.Masks.Add();
            _objMask.Shapes.AddByPointsArray(input.Length / 3, input, true, EMaskEngineUnit.emeUnitMils);

            Mask _finalMask = null;
            if ( inShapes != null )
                _finalMask = _me.BooleanOp(EMaskEngineBooleanOp.emeBooleanOpAND, _objMask, _inShapeMask);
            else
                _finalMask = _objMask;
            if ( outShapes != null )
                _finalMask = _me.BooleanOp(EMaskEngineBooleanOp.emeBooleanOpSubtract, _finalMask, _outShapeMask);

            List<object[,]> _returnList = new List<object[,]>();
            foreach ( Shape _s in _finalMask.Shapes )
                _returnList.Add((object[,])_s.get_PointsArray(EMaskEngineUnit.emeUnitMils));

            return _returnList;
        }

        public static object[,] V2_Transform( object[,] input, double dX, double dY )
        {
            for ( int i = 0;i < input.Length / 3;i++ )
            {
                input[0, i] = (double)input[0, i] + dX;
                input[1, i] = (double)input[1, i] + dY;
            }
            return input;
        }

    }
}

