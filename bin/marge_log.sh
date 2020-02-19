#!/bin/bash
set -u

cat $1/*.log > $1/marge_log@$1.csv

if [ -z `which nkf` ]; 
  brew install nkf
  nkf -s --overwrite $1/marge_log@$1.csv
else
  nkf -s --overwrite $1/marge_log@$1.csv
fi
