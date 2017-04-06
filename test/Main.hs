{-# LANGUAGE OverloadedStrings #-}

module Main (main)
  where
import           Test.SmallCheck.Series (Positive (..))
import           Test.Tasty             (defaultMain, testGroup)
import           Test.Tasty.HUnit       (testCase)
import           Test.Tasty.HUnit       ((@=?))
import           Test.Tasty.SmallCheck  (testProperty)
import           WriteXLSX
import           WriteXLSX.ExtractKeys
import ReadXLSX
import           Data.ByteString.Lazy (ByteString)


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
        json @=? "{\"Col1\":[\"\195\169\",\"\194\181\",\"b\"],\"Col2\":[\"2017-01-25\",\"2017-01-26\",\"2017-01-27\"],\"\194\181\":[1,1,1]}"
    ]

testExtractKeys :: [String]
testExtractKeys = extractKeys "{\"a\":2,\"\\u00b5\":1,\"µ\":2}"

testRXL1 :: IO ByteString
testRXL1 = sheetToJSON "test/simpleExcel.xlsx" "Sheet1" "data" True True Nothing Nothing

testRXL2 :: IO ByteString
testRXL2 = sheetToJSON "test/utf8.xlsx" "Sheet1" "data" True True Nothing Nothing
