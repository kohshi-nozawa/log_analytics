#!/bin/bash
set -u

date=`date '+%Y%m%d'`

for file in `ls $date`
do
  name=`ls $date/$file | cut -c 10-`
  while read line
  do
    echo $name"|"$line >> $date/marge$date.log
  done < $date/$file
done

if [ -z `which nkf` ]; 
  brew install nkf
  nkf -s --overwrite $date/marge$date.log
else
  nkf -s --overwrite $date/marge$date.log
fi