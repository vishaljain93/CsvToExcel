#!/bin/bash

inputfilepath=$1
outputfilepath=$2
newline=""
pipe="|"
filenumber=1
recordcount=1
RED=`tput setaf 1`
GREEN=`tput setaf 2`

echo "This script will convert csv files to excel files."
echo "It will create multiple files and each file will contain 50K records."

if [ $# -ne 2 ]; then
	echo -e "\n\t\t\t\t${RED}Please provide input and output file path respectively including file names without file extension."
	exit 1
fi
while read line; do
	columnnumber=$(( `cat ${inputfilepath} | head -1 | grep -o "," | wc -l` + 1 ))
	newline=`echo ${line} | cut -d"," -f1`
	
	for ((i = 2; i <= $columnnumber; i++)); do
        	newline=${newline}${pipe}`echo ${line} | cut -d"," -f$i`
	done
	echo $newline >> ${outputfilepath}$filenumber.xlsx
	newline=""
	recordcount=$(( $recordcount + 1 ))
	if [[  $recordcount -gt 3000 ]]; then
		echo "${filenumber} bunch of 50K records has been completed."
		recordcount=1
		filenumber=$(( $filenumber + 1 ))
		sleep 3
	fi
done < ${inputfilepath}

echo "${GREEN}Files have been created successfully!!!"
echo "Each file contain 50K recordss."
echo "Use | (pipe) as a separator to open the file."
