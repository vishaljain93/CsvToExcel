Requirements:
    -java: 1.8
    - input path
    -output path
    
1. Please provide input and output path as command line arguments. Please refer below example:

java -jar target/CsvToXml-1.0-SNAPSHOT-jar-with-dependencies.jar input_path output_path

Format for path variables:
input_path = "D:\\Projects\\filename.csv"
output_path = "D:\\Projects\\filename.xls"

java -jar target/CsvToXml-1.0-SNAPSHOT-jar-with-dependencies.jar "D:\\Projects\\filename.csv" "D:\\Projects\\filename.xls"

A sample csv file of 10000 records is in src/main/resources folder.

Note: Do not change the jar name and input or output extension.

To run the project using IDE, please find IDESetup.md