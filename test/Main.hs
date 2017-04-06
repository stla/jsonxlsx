{-# LANGUAGE OverloadedStrings #-}

module Main (main)
  where
import           Data.ByteString.Lazy  (ByteString)
import           Data.ByteString.Lazy.Internal (unpackChars)
import           ReadXLSX
import           System.Directory
import           Test.Tasty            (defaultMain, testGroup)
import           Test.Tasty.HUnit      (testCase)
import           Test.Tasty.HUnit      ((@=?))
import           Test.Tasty.SmallCheck (testProperty)
import           WriteXLSX
import           WriteXLSX.ExtractKeys


main :: IO ()
main = defaultMain $
  testGroup "Tests"
    [ testCase "extractKeys" $
        testExtractKeys @=? ["a", "\181", "\181"],
      testCase "read xl1" $ do
        json <- testRXL1
        json @=? "{\"A\":[1]}",
      testCase "read xl2" $ do
        json <- testRXL2
        json @=? "{\"Col1\":[\"\195\169\",\"\194\181\",\"b\"],\"Col2\":[\"2017-01-25\",\"2017-01-26\",\"2017-01-27\"],\"\194\181\":[1,1,1]}",
      testCase "write and read" $ do
        json <- testWriteAndRead
        json @=? True -- "{\"Col1\":[1,2]}"
    ]

testExtractKeys :: [String]
testExtractKeys = extractKeys "{\"a\":2,\"\\u00b5\":1,\"µ\":2}"

testRXL1 :: IO ByteString
testRXL1 = sheetToJSON "test/simpleExcel.xlsx" "Sheet1" "data" True True Nothing Nothing

testRXL2 :: IO ByteString
testRXL2 = sheetToJSON "test/utf8.xlsx" "Sheet1" "data" True True Nothing Nothing

testWriteAndRead :: IO Bool
testWriteAndRead = do
  let json = "{\"Col1\":[1,2]}"
  tmpDir <- getTemporaryDirectory
  let tmpFile = tmpDir ++ "/xlsx.xlsx"
  write <- write1 json True tmpFile False
  jjson <- sheetToJSON tmpFile "Sheet1" "data" True True Nothing Nothing
  return $ json == (unpackChars jjson)
