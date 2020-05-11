The shell script is to convert csv file to excel file.

It takes two command line arguments:
1. Input file path.
2. Output file path.

It gives an error message if arguments are less or greater than 2.

Note: Do not provide file extension in output file path.

For large number of records it creates multiple excel files.
Each file will contain 50K records.

Output file name will contain the number of order of its creation.

For example:
output1.xlsx
output2.xlsx and so on.

To execute the file, please follow below steps:
1. chmod +x csvtoexcel.sh
2. ./csvtoexcel.sh input_file_path output_file_path

Note: Newly generated files are not getting open in windows as linux is not able to provide the format Microsoft Excel asks to handle the file.
