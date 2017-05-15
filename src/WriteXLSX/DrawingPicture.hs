{-# LANGUAGE OverloadedStrings #-}
module WriteXLSX.DrawingPicture
  where
import           Control.Lens
import Data.ByteString.Lazy (ByteString)
import           Codec.Xlsx

defaultShapeProperties :: ShapeProperties
defaultShapeProperties =
  ShapeProperties {
                    _spXfrm     = Nothing,
                    _spGeometry = Just PresetGeometry,
                    _spFill     = Nothing,
                    _spOutline  = Nothing
                  }

drawingPicture :: ByteString -> Drawing
drawingPicture image = Drawing {_xdrAnchors = [set anchObject pic anchor]}
  where anchor = simpleAnchorXY (2, 3) (positiveSize2D 3000000 3000000) $
                  picture DrawingElementId{unDrawingElementId = 1} fileInfo
        pic = set picShapeProperties defaultShapeProperties (_anchObject anchor)
        fileInfo = FileInfo
                       {
                         _fiFilename = "image1.png",
                         _fiContentType = "image/png",
                         _fiContents =  image
                       }
