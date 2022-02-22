#! /bin/bash

pushd .
cd /home/cy/Ding-Notify/
python3 ./main.py $1 2> ./py_output

popd
