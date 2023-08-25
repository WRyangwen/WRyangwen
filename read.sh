#! /bin/bash

result_log=./result
if [[ ! -e  ${result_log} ]]
then
   mkdir ${result_log}
fi

fio ./128K_seq_r.fio --output=${result_log}/read.log  --aux-path=${result_log}

if [[ $? == 0 ]]
then
    echo "Pass !!!!"
else
    echo "Failed !!!!"
fi
