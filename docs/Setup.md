Requirements:
    -java: 1.8
    -input path
    -output path
    
Project Root Directory: CsvToExcel/
    
To create the jar file, execute following command from project root directory: mvn clean install

jar file will be in the CsvToExcel/target folder.

CsvToExcel-1.0-SNAPSHOT-jar-with-dependencies.jar will have all the related dependencies within it.

Please follow below steps to run the program using jar file.
1. Go to project root directory
2. Please provide input and output path as command line arguments.
3. Run below command from the project root directory.

java -jar target/CsvToExcel-1.0-SNAPSHOT-jar-with-dependencies.jar input_path output_path

For Windows=>
input_path = "D:\\Projects\\filename.csv"
output_path = "D:\\Projects\\filename.xlsx"

java -jar target/CsvToExcel-1.0-SNAPSHOT-jar-with-dependencies.jar "D:\\Projects\\filename.csv" "D:\\Projects\\filename.xlsx"

For Linux=>
input_path = "~/Desktop/filename.csv"
output_path = "~/Desktop/filename.xlsx"

java -jar target/CsvToExcel-1.0-SNAPSHOT-jar-with-dependencies.jar "~/Desktop/filename.csv" "~/Desktop/filename.xlsx"

A sample csv file of 10000 records is in src/main/resources folder.

Note: Do not change the jar name and input or output extension.

To run the project using IDE, please find docs/IDESetup.md