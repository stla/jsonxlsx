{-# LANGUAGE OverloadedStrings #-}
module WriteXLSX.DrawingPicture
  where
import           Codec.Xlsx
import           Control.Lens
import           Data.ByteString.Lazy (ByteString)

defaultShapeProperties :: ShapeProperties
defaultShapeProperties =
  ShapeProperties {
                    _spXfrm     = Nothing,
                    _spGeometry = Just PresetGeometry,
                    _spFill     = Nothing,
                    _spOutline  = Nothing
                  }

drawingPicture :: ByteString -> (Int, Int, Int, Int) -> Drawing
drawingPicture image coordinates =
  Drawing {_xdrAnchors = [set anchObject pic anchor]}
  where anchor = simpleAnchorXY (left-1, top-1) (positiveSize2D cx cy) $
                  picture DrawingElementId{unDrawingElementId = 1} fileInfo
        pic = set picShapeProperties defaultShapeProperties (_anchObject anchor)
        fileInfo = FileInfo
                       {
                         _fiFilename = "image1.png",
                         _fiContentType = "image/png",
                         _fiContents =  image
                       }
        (top, left, width, height) = coordinates
        cx = toInteger $ 9525*width
        cy = toInteger $ 9525*height
